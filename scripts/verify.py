"""Independently reconcile a recalculated Excel readback to the source record."""
import argparse
from datetime import date, timedelta
import json
import math
import sys

from reconcile import read_json, reconcile


def month_start_after(day, months):
    offset = day.year * 12 + day.month - 1 + months
    return date(offset // 12, offset % 12 + 1, 1)


def verify(deal, book):
    source = reconcile(deal)
    failures = source["differences"].copy()
    follow_up = source["open"].copy()
    checked = []
    assumptions = {k: d["value"] for k, d in deal["decisions"].items()}

    def compare(label, actual, expected, tolerance=0.02):
        if expected is None:
            follow_up.append(f"{label}: independent control missing")
        elif type(actual) not in (int, float) or abs(actual - expected) > tolerance:
            failures.append(f"{label}: Excel {actual!r}; expected {expected}")
        else:
            checked.append(label)

    def require(label, ok):
        (checked if ok else failures).append(label)

    def number(k):
        v = book[k]
        if type(v) not in (int, float) or not math.isfinite(v):
            raise ValueError(f"{k} has no finite recalculated value")
        return v

    require("Correct subject", str(book["name"]).strip().casefold() == deal["property"].strip().casefold())
    for field, key, tol in [
            ("price", "purchase_price", .01), ("units", "acquired_units", 0),
            ("rsf", "acquired_rsf", .01), ("hold", "hold_months", 0),
            ("gp_share", "gp_share", 1e-8), ("exit_cap", "exit_cap", 1e-8),
            ("sale_cost_rate", "sale_cost_rate", 1e-8), ("renewal_rate", "renewal_rate", 1e-8),
            ("current_leasing", "current_leasing", .0001), ("forecast_leasing", "forecast_leasing", 1e-8),
            ("reserve_per_unit","annual_reserve_per_unit",.01)]:
        compare(key, book[field], assumptions[key], tol)
    if assumptions["close_date"] is not None:
        require("Acquisition date", book["close"] == assumptions["close_date"])
    opened = date.fromisoformat(book["close"])
    hold = number("hold")
    if hold < 1 or int(hold) != hold:
        raise ValueError("hold must be positive whole months")
    hold = int(hold)
    if opened.day == 1:
        require("Last day of modeled hold", date.fromisoformat(book["sale"]) == month_start_after(opened,hold)-timedelta(days=1))
    else:
        follow_up.append("Review stub period and contractual closing date")
    if deal["t12"]:
        compare("Seller expense reconciliation", book["history_expenses"], source["history"]["total_expenses"])
        start = date.fromisoformat(deal["t12"]["start"])
        require("Seller history period", book["history_end"] == (month_start_after(start,12)-timedelta(days=1)).isoformat())
    if assumptions.get("history_income_complete") is False:
        require("Incomplete historical NOI shown as unavailable", book["history_noi"] in ("", None, "N/A"))

    price = number("price")
    cost_rate = assumptions["buyer_cost_rate"]
    compare("Buyer closing costs", book["buyer_costs"], price*cost_rate if cost_rate is not None else None)
    compare("Sources equal uses", book["sources"], number("uses"))
    compare("GP and LP checks", book["gp"], number("equity")*number("gp_share"))
    compare("Partner capital totals", book["lp"], number("equity")-number("gp"))
    compare("Opening MFFO equity basis", book["mffo_open"], number("mffo_close_at_purchase"))
    compare("Operating NOI", book["noi"], number("egi")-number("opex"))
    compare("Independent Year 1 NOI", book["noi"], assumptions.get("expected_year1_noi"), 1.00)
    if price <= 0:
        raise ValueError("Purchase price must be positive")
    compare("Year 1 cap rate", book["cap_rate"], number("noi")/price, .00001)
    if assumptions.get("retail_area", None) == 0 and number("exit_cap") > 0:
        compare("Residential residual value", book["gross_sale"], number("exit_noi")/number("exit_cap"), 1)
    else:
        follow_up.append("Tie each component of residual value to its NOI and exit cap")

    dates, cf = book["dates"], book["levered_cash"]
    if len(cf) < hold+1 or len(dates) != len(cf):
        raise ValueError("Cash flow grid does not cover the stated hold or date count")
    cf_hold = cf[:hold+1]
    if any(type(v) not in (float,int) for v in cf_hold) or cf_hold[0] >= 0:
        raise ValueError("Expected a numeric dated cash-flow series with opening equity outflow")
    require("No additional cash flows after sale", all(v in (0, "", None) for v in cf[hold+1:]))
    days = [date.fromisoformat(d) for d in dates[:hold+1]]
    require("Monthly cash-flow periods", all(day.replace(day=1)==month_start_after(days[0],i) for i,day in enumerate(days)))
    require("Exit cash flow in sale month", days[-1].replace(day=1)==date.fromisoformat(book["sale"]).replace(day=1))
    positive = sum(max(v,0) for v in cf_hold)
    invested = -sum(min(v,0) for v in cf_hold)
    if not any(v < 0 for v in cf_hold[1:]):
        compare("Equity multiple from cash flows", book["multiple"], positive/invested, .00001)
    else:
        follow_up.append("Reconcile equity multiple from gross contributions and distributions by period")
    rate = number("irr")
    if rate <= -1 or positive <= 0:
        failures.append("A valid XIRR requires positive and negative cash flows and a rate above -100%")
    else:
        npv = sum(v/(1+rate)**((day-days[0]).days/365) for v,day in zip(cf_hold, days))
        compare("Dated cash-flow NPV at displayed XIRR", npv, 0, max(1, invested*1e-6))
        if any(v < 0 for v in cf_hold[1:]):
            follow_up.append("Later equity calls: check IRR root and gross contributions/distributions for the multiple")

    loan = assumptions["loan_structure"]
    if isinstance(loan,dict) and loan.get("kind") == "fixed_io_at_close":
        # Applies to a single acquisition loan, no future funding, 30/360 and payoff at sale.
        require("Screened loan covers the hold", loan["term_months"] >= hold)
        principal = loan["amount"] if "amount" in loan else price * loan["ltv"]
        fee = principal * loan["fee_rate"]
        interest = principal * loan["rate"]
        require("Positive fixed IO loan and coupon", principal>0 and interest>0)
        for key, expected in [("opening_debt",principal),("initial_principal",principal),
                              ("debt",principal),("loan_costs",fee),("opening_fee_cf",-fee),
                              ("year1_debt_service",-interest*min(hold,12)/12)]:
            compare(key, book[key], expected)
        if hold >= 12 and interest>0:
            compare("Interest coverage", book["loan1_dscr"],number("noi")/interest,.001)
            compare("Debt yield", book["loan1_yield"],number("noi")/principal,.0001)
        payoffs = book["loan1_payoff"]
        if len(payoffs) < hold+1:
            raise ValueError("Loan schedule ends before disposition")
        compare("Balloon paid at sale", payoffs[hold], -principal)
        require("No earlier or duplicate balloon", all(v in (0,None,"") for i,v in enumerate(payoffs) if i!=hold))
        compare("Initial equity and debt fund uses", -cf_hold[0]+principal,number("uses"))
        if hold >= 12:
            compare("Year 1 cash-on-cash", book["cash_on_cash"], number("year1_equity_cf")/-cf_hold[0], .00001)
    elif isinstance(loan,dict) and loan.get("kind") == "cash":
        compare("Cash purchase has no loan", book["debt"],0)
    else:
        follow_up.append("Trace each actual facility, draws, amortization, refi and payoff in Excel")
    errors = book["formula_errors"]
    if not isinstance(errors,list):
        raise ValueError("Expected workbook-wide formula error inventory")
    follow_up.extend(f"Classify formula error and trace active-case dependencies: {err}" for err in errors)
    return {"status":"differences" if failures else ("review_items" if follow_up else "checked"),
            "passed":checked,"differences":failures,"open":follow_up}


if __name__ == "__main__":
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument("deal")
    p.add_argument("readback")
    a=p.parse_args()
    try:
        result=verify(read_json(a.deal),read_json(a.readback))
        print(json.dumps(result,indent=2,allow_nan=False))
        sys.exit(int(bool(result["differences"])))
    except (ValueError, KeyError, TypeError, OSError, ArithmeticError) as e:
        sys.exit(f"Could not verify workbook: {e}")

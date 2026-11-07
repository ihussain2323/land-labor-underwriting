"""Calculate documented gap allowances, dated tax bills and incremental surcharge cases."""
import argparse
from calendar import monthrange
from datetime import date
import json
import sys
from reconcile import read_json, validate
from verify import month_start_after


def amount_in_month(start, end, annual, month):
    following = month_start_after(month, 1)
    days = max(0, (min(end, following)-max(start, month)).days)
    return annual * days / monthrange(month.year, month.month)[1] / 12


def spans(entries):
    result = []
    for item in entries:
        start = date.fromisoformat(item["from"])
        end = date.fromisoformat(item["until"])
        if end <= start:
            raise ValueError("Every period needs an exclusive end date after its start")
        result.append((start, end, item))
    return sorted(result, key=lambda row: row[0])


def analyze(deal):
    validate(deal)
    d = {k:v["value"] for k,v in deal["decisions"].items()}
    start = date.fromisoformat(d["close_date"])
    if start.day != 1:
        raise ValueError("For a stub acquisition period, use actual Excel period boundaries")
    # Include the forward year for disposition analysis.
    months = [month_start_after(start,i) for i in range(d["hold_months"]+12)]
    result = {"property":deal["property"], "months":[m.isoformat() for m in months]}
    if d.get("expense_gaps") is not None:
        gaps = {}
        for line in d["expense_gaps"]:
            k=line["category"]
            if k in gaps:
                raise ValueError(f"Duplicate gap category: {k}")
            allowance=line["benchmark_per_unit"]*d["acquired_units"]-line["already_in_seller"]
            if allowance < 0:
                raise ValueError(f"{k}: overlapping seller charges exceed benchmark; resolve allocation")
            gaps[k]=allowance
        result["annual_gap_allowances"]=gaps
    tax=d.get("tax_schedule")
    if isinstance(tax,dict) and "bills" in tax:
        bills=spans(tax["bills"])
        for prev,nxt in zip(bills,bills[1:]):
            if prev[1] != nxt[0]:
                raise ValueError("Tax bills must be consecutive and must not overlap")
        if not bills or bills[0][0]>start or bills[-1][1]<month_start_after(months[-1],1):
            raise ValueError("Supply sourced or explicitly modeled taxes through the forward year")
        result["monthly_taxes"]=[sum(amount_in_month(a,b,entry["annual"],m) for a,b,entry in bills) for m in months]
    if d.get("surcharge_cases") is not None:
        cases={}
        for case in d["surcharge_cases"]:
            name=case["name"]
            if name in cases:
                raise ValueError(f"Duplicate surcharge case: {name}")
            collected=case["collection_fraction"]
            management=case["management_rate"]
            if not all(0<=x<=1 for x in (collected,management)):
                raise ValueError(f"{name}: collection and management must be fractions")
            periods=spans(case["periods"])
            if any(prev[1]>nxt[0] for prev,nxt in zip(periods,periods[1:])):
                raise ValueError(f"{name}: consolidate eligible cohorts into nonoverlapping periods")
            if any(e["eligible_annual_rent"]<0 or not 0<=e["rate"]<=1 for _,_,e in periods):
                raise ValueError(f"{name}: invalid eligible rent or rate")
            gross=[sum(amount_in_month(a,b,e["eligible_annual_rent"]*e["rate"],m)
                       for a,b,e in periods) for m in months]
            cases[name]={"gross_monthly":gross,"incremental_noi_monthly":[v*collected*(1-management) for v in gross]}
        result["surcharge_cases"]=cases
    if len(result)==2:
        raise ValueError("No structured gap, tax-bill or surcharge-case inputs found")
    return result


if __name__ == "__main__":
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("deal")
    try:
        print(json.dumps(analyze(read_json(parser.parse_args().deal)),indent=2,allow_nan=False))
    except (TypeError,ValueError,KeyError,OSError) as e:
        sys.exit(str(e))

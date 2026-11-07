"""Emit evidence-linked candidate writes for verified template slots. Does not edit Excel."""
import argparse
from datetime import date, timedelta
import json
from pathlib import Path
import sys
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import range_boundaries

from reconcile import read_json, reconcile
from verify import month_start_after


def plan(deal, inspected):
    check = reconcile(deal)
    if check["differences"]:
        raise ValueError("; ".join(check["differences"]))
    decisions = deal["decisions"]
    writes, manual, handled = [], [], set()

    def bind(sheet, field, values, key, evidence):
        mapping = inspected.get(f"{sheet} / {field}")
        if not mapping:
            manual.append(f"{sheet}: resolve {field} by its live label")
            return
        c,r,c2,r2 = range_boundaries(mapping["address"])
        cells = [f"{get_column_letter(col)}{row}" for row in range(r,r2+1) for col in range(c,c2+1)]
        if not isinstance(values,list):
            values = [values]*len(cells)
        if len(values)!=len(cells) or set(cells).intersection(mapping["preserve_formulas"]):
            manual.append(f"{sheet}!{mapping['address']}: protected formula or incompatible dimensions; trace {key}")
            return
        for addr,value in zip(cells,values):
            if value is not None:
                writes.append({"sheet":sheet,"cell":addr,"value":value,"field":key,"source":evidence})
        handled.add(key)

    tx,asset="Transaction Inputs","Asset Inputs"
    fields = {tx: {"purchase_price":"purchase_price","buyer_cost_rate":"buyer_cost_rate",
                   "close_date":"close_date","hold_months":"hold_months","exit_cap":"exit_cap",
                   "sale_cost_rate":"sale_cost_rate","gp_share":"gp_share"},
              asset:{"current_leasing":"current_leasing","renewal_rate":"renewal_rate",
                     "annual_reserve_per_unit":"annual_reserve_per_unit","retail_area":"retail_area"}}
    for sheet, pairs in fields.items():
        for key,slot in pairs.items():
            d=decisions.get(key)
            if d and d["value"] is not None and d["status"]!="scenario":
                bind(sheet,slot,d["value"],key,d["source"])
    def active(k):
        return decisions[k]["value"] if decisions[k]["status"]!="scenario" else None
    occ = active("forecast_leasing")
    if occ is not None:
        mapping=inspected.get(asset+" / leasing")
        if mapping:
            c,_,end,_=range_boundaries(mapping["address"])
            bind(asset,"leasing",[active("current_leasing")]+[occ]*(end-c),"forecast_leasing",decisions["forecast_leasing"]["source"])
    if deal["t12"]:
        end=month_start_after(date.fromisoformat(deal["t12"]["start"]),1)-timedelta(days=1)
        bind("T-12","first_month",end.isoformat(),"t12",deal["t12"]["source"])
        aliases=decisions.get("history_accounts",{}).get("value") or {}
        used=set()
        for category,monthly in deal["t12"]["expenses"].items():
            account=aliases.get(category,category)
            if account in used:
                raise ValueError(f"Consolidate seller categories mapped to {account} before proposing writes")
            used.add(account)
            for role,value in [("label",category),("account",account),("monthly",monthly)]:
                bind("T-12",f"history.{account}.{role}",value,"t12",deal["t12"]["source"])
    operating=decisions.get("operating_lines")
    if operating and isinstance(operating["value"],list) and operating["status"]!="scenario":
        for line in operating["value"]:
            category, basis = line["account"], line["selector"]
            if basis not in ("$ / Unit","% of EGR"):
                manual.append(f"Resolve operating selector for {category}: {basis}")
                continue
            slot="per_unit" if basis=="$ / Unit" else "rate"
            bind(asset,f"operating.{category}.selector",basis,"operating_lines",operating["source"])
            bind(asset,f"operating.{category}.{slot}",line["input"],"operating_lines",operating["source"])
            if "growth" in line:
                bind(asset,f"operating.{category}.growth",line["growth"],"operating_lines",operating["source"])
    # The master has four unit slots. Preserve existing type order on refresh by mapping manually.
    mix=check.get("unit_mix",{})
    order=decisions.get("unit_type_order",{}).get("value")
    valid_order = bool(mix) and isinstance(order,list) and len(order)<=4 and len(order)==len(set(order)) and set(order)==set(mix)
    if valid_order:
        pad=lambda vals: vals+[None]*(4-len(vals))
        for slot,vs in [('unit_types',order),('unit_counts',[mix[k]['count'] for k in order]),('unit_sizes',[mix[k]['avg_rsf'] for k in order])]:
            bind(asset,slot,pad(vs),"unit_mix",deal["roll"]["source"])
        rate=active("market_rent_growth")
        if rate is not None:
            for i in range(1,len(order)+1):
                bind(asset,f"rent_growth_type{i}",rate,"market_rent_growth",decisions["market_rent_growth"]["source"])
    elif mix:
        manual.append("Confirm unit-type grouping/order, allocated SF and unused slots before writing unit mix")
    loan=active("loan_structure")
    if isinstance(loan,dict) and loan.get("kind") in ("cash","fixed_io_at_close"):
        key="loan_structure"
        source=decisions[key]["source"]
        single=loan["kind"]=="fixed_io_at_close"
        bind(tx,"loan_active",["No","No","Yes" if single else "No"],key,source)
        if single:
            term=loan.get("term_months")
            if not term:
                manual.append("Confirm or select an explicit screening term for Loan 1")
            settings={"loan_repayment_source":"Sale","loan_funding_month":0,"loan_term":term,
                      "loan_io":term,"loan_amortization":0,"loan_rate_type":"Fixed Rate",
                      "loan_day_count":"30/360","loan_rate":loan["rate"],
                      "loan_method":"Amount" if "amount" in loan else "Purchase LTV",
                      "loan_size_or_percent":loan.get("amount",loan.get("ltv")),"loan_fee":loan["fee_rate"]}
            for slot,v in settings.items():
                if v is not None:
                    bind(tx,slot,[None,None,v],key,source)
    # Explicitly sourced type rates allow regulated collectible rents and unrestricted
    # comparables to coexist. Never silently substitute one for the other.
    base=decisions.get("starting_rent_by_type")
    if base and isinstance(base["value"],dict) and base["status"]!="scenario" and valid_order:
        for i,typename in enumerate(order,1):
            if typename in base["value"]:
                bind(asset,f"starting_rent{i}",base["value"][typename],"starting_rent_by_type",base["source"])
            else:
                manual.append(f"Confirm starting rent for {typename}")
    for key,field in [("research_rent","research_rent"),("research_value","research_value")]:
        entry=decisions.get(key)
        if entry and isinstance(entry["value"],dict) and entry["status"]!="scenario":
            series=entry["value"]
            slot=inspected.get(asset+" / "+field)
            if slot and "latest_yoy" in series and "mean_yoy" in series:
                c,_,c2,_=range_boundaries(slot["address"])
                bind(asset,field,[series["latest_yoy"]]+[series["mean_yoy"]]*(c2-c),key,entry["source"])
    for k,d in decisions.items():
        if k not in handled:
            manual.append(f"{k} ({d['status']}): {d['destination']}; {d['basis']}")
    seen = set()
    for w in writes:
        cell=(w["sheet"],w["cell"])
        if cell in seen:
            raise ValueError(f"Conflicting writes for {cell}")
        seen.add(cell)
    return {"property":deal["property"],"candidate_writes":writes,"resolve_in_excel":manual,"open":check["open"]}


if __name__ == "__main__":
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('deal')
    p.add_argument('inspected')
    a=p.parse_args()
    try:
        print(json.dumps(plan(read_json(a.deal),read_json(a.inspected)),indent=2,allow_nan=False))
    except (ValueError,KeyError,TypeError,OSError) as e:
        sys.exit(str(e))

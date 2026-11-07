"""Compare approved candidate writes to the actual saved workbook, including formulas and dates."""
import argparse
from datetime import date, datetime
import json
import math
import sys

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_to_tuple
from openpyxl.utils.datetime import from_excel

from reconcile import read_json


def check(path, plan):
    grouped = {}
    for entry in plan["candidate_writes"]:
        sheet = grouped.setdefault(entry["sheet"], {})
        address = entry["cell"].upper()
        if address in sheet:
            raise ValueError(f"Repeated target {entry['sheet']}!{address}")
        sheet[address] = entry["value"]
    book = load_workbook(path, read_only=True, data_only=False)
    differences = []
    count = 0
    try:
        for name, expected in grouped.items():
            cells = {coordinate_to_tuple(addr):(addr,v) for addr,v in expected.items()}
            rows = book[name].iter_rows(max_row=max(r for r,c in cells),max_col=max(c for r,c in cells))
            for r,row in enumerate(rows,1):
                for c,cell in enumerate(row,1):
                    if (r,c) not in cells:
                        continue
                    addr,want = cells[(r,c)]
                    got = cell.value
                    if isinstance(want,str):
                        try:
                            when = date.fromisoformat(want)
                        except ValueError:
                            when = None
                        if when is not None:
                            if type(got) in (int,float):
                                got = from_excel(got,book.epoch)
                            if isinstance(got,datetime):
                                got = got.date()
                            if isinstance(got,date):
                                got=got.isoformat()
                    if type(want) in (int,float):
                        same=type(got) in (int,float) and math.isclose(want,got,rel_tol=1e-10,abs_tol=1e-9)
                    else:
                        same=want==got
                    if not same:
                        differences.append({"sheet":name,"cell":addr,"expected":want,"actual":got})
                    count+=1
    finally:
        book.close()
    if count != len(plan["candidate_writes"]):
        raise ValueError("Readback did not cover every proposed cell")
    return {"checked":count,"differences":differences}


if __name__=="__main__":
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument("xlsx")
    parser.add_argument("approved_plan")
    argv=parser.parse_args()
    try:
        result=check(argv.xlsx,read_json(argv.approved_plan))
        print(json.dumps(result,indent=2,default=str))
        sys.exit(int(bool(result["differences"])))
    except (KeyError,TypeError,ValueError,OSError) as e:
        sys.exit(str(e))

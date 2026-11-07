"""Inspect mapped inputs or export cached Excel outputs. This utility only reads files."""
import argparse
from datetime import date, datetime, timedelta
import json
from pathlib import Path
import sys
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries

ROOT = Path(__file__).resolve().parents[1]


def normalize(s):
    return " ".join(str(s).split()).casefold()


def inspect(path, mapping, capture):
    # Formula and value views are intentionally separate. Cached values require Excel recalc + save.
    book = load_workbook(path, data_only=capture, read_only=True)
    output = {}
    try:
        for name, spec in mapping.items():
            if name not in book.sheetnames:
                raise ValueError(f"Missing {name}; resolve a deal-specific map")
            addresses = spec["outputs" if capture else "inputs"]
            bounds = {a: range_boundaries(a) for a in [*spec["anchors"], *addresses.values()]}
            max_row = max(b[3] for b in bounds.values())
            max_col = max(b[2] for b in bounds.values())
            rows = list(book[name].iter_rows(max_row=max_row, max_col=max_col))
            for addr, label in spec["anchors"].items():
                c, r, _, _ = bounds[addr]
                allowed = label if isinstance(label,list) else [label]
                if normalize(rows[r-1][c-1].value) not in {normalize(v) for v in allowed}:
                    raise ValueError(f"{name}!{addr}: expected {label!r}; remap this layout before use")
            for key, addr in addresses.items():
                c, r, c2, r2 = bounds[addr]
                cells = [cell for row in rows[r-1:r2] for cell in row[c-1:c2]]
                if capture:
                    vs = [cell.value for cell in cells]
                    vs = [v.date().isoformat() if isinstance(v, datetime) else
                          v.isoformat() if isinstance(v, date) else v for v in vs]
                    if key in ('dates','sale','close','history_end'):
                        # Match the file's own Excel date system even with General-formatted date cells.
                        vs = [(book.epoch + timedelta(days=v)).date().isoformat() if type(v) in (int,float) else v for v in vs]
                    output[key] = vs[0] if len(vs)==1 else vs
                else:
                    output[f"{name} / {key}"] = {"address":addr, "preserve_formulas":[cell.coordinate for cell in cells if cell.data_type=="f"]}
        if capture:
            output["formula_errors"] = [
                f"{sheet.title}!{cell.coordinate}"
                for sheet in book for row in sheet.iter_rows() for cell in row
                if cell.data_type == "e"
            ]
    finally:
        book.close()
    return output


if __name__ == "__main__":
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument("action", choices=["inspect","capture"])
    p.add_argument("xlsx")
    p.add_argument("--map", default=str(ROOT/"assets/master-map.json"))
    a=p.parse_args()
    try:
        result=inspect(a.xlsx,json.loads(Path(a.map).read_text()),a.action=="capture")
        print(json.dumps(result,indent=2,allow_nan=False,default=str))
    except (ValueError, KeyError, OSError) as e:
        sys.exit(str(e))

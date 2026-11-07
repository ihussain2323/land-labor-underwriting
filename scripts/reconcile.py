"""Validate a deal record and reconcile unit-level and monthly source data."""
import argparse
from collections import defaultdict
from datetime import date
import json
import math
from pathlib import Path
import sys

from jsonschema import Draft202012Validator, FormatChecker

ROOT = Path(__file__).resolve().parents[1]
RATES = {"gp_share", "current_leasing", "forecast_leasing", "renewal_rate",
         "buyer_cost_rate", "exit_cap", "sale_cost_rate"}


def read_json(path):
    def reject(value):
        raise ValueError(f"Invalid JSON number: {value}")
    return json.loads(Path(path).read_text(encoding="utf-8"), parse_constant=reject)


def validate(deal):
    schema = read_json(ROOT / "schemas/deal.schema.json")
    errors = Draft202012Validator(schema, format_checker=FormatChecker()).iter_errors(deal)
    messages = [f"{e.json_path}: {e.message}" for e in errors]
    if messages:
        raise ValueError("\n".join(messages))
    for key, item in deal["decisions"].items():
        if item["source"] not in deal["sources"]:
            raise ValueError(f"{key}: unknown source ID {item['source']}")
        v = item["value"]
        if v is None:
            continue
        if key in RATES and (type(v) not in (int, float) or not 0 <= v <= 1):
            raise ValueError(f"{key}: expected a decimal fraction in [0,1]")
        if key in {"purchase_price", "acquired_rsf", "exit_cap"} and (
                type(v) not in (int, float) or v <= 0):
            raise ValueError(f"{key}: expected a positive number")
        if key in {"hold_months", "acquired_units"} and (type(v) is not int or v < 1):
            raise ValueError(f"{key}: expected a positive integer")
        if key == "close_date":
            date.fromisoformat(str(v))
    for key in ("roll", "t12"):
        section = deal[key]
        if section and section["source"] not in deal["sources"]:
            raise ValueError(f"{key}: unknown source ID")


def reconcile(deal):
    validate(deal)
    errors = []
    pending = [key for key, d in deal["decisions"].items() if d["status"] == "unknown"]
    pending.extend(item["item"] for item in deal["issues"])
    result = {"property": deal["property"], "differences": errors, "open": pending}
    d = {k: v["value"] for k, v in deal["decisions"].items()}
    if deal["roll"]:
        units = deal["roll"]["units"]
        seen = set()
        groups = defaultdict(list)
        for u in units:
            uid = u["id"].strip().casefold()
            if not uid or uid in seen:
                errors.append(f"Blank or duplicate unit ID {u['id']!r}")
            seen.add(uid)
            groups[u["type"]].append(u)
        def total(rows, key):
            vals = [u[key] for u in rows]
            return sum(vals) if all(v is not None for v in vals) else None
        def group(rows):
            area = total(rows, "rsf")
            sums = {k: total(rows, k) for k in ("contract", "legal", "market")}
            return {"count": len(rows), "rsf": area,
                    "avg_rsf": area / len(rows) if area else None,
                    "annual_scheduled": {k: v * 12 if v is not None else None for k, v in sums.items()},
                    "monthly_per_rsf": {k: v / area if area and v is not None else None for k, v in sums.items()}}
        result["unit_mix"] = {k: group(v) for k, v in groups.items()}
        result["residential"] = group(units)
        for key, value in (("acquired_units", len(units)), ("acquired_rsf", result["residential"]["rsf"])):
            if value is None:
                pending.append(f"Incomplete unit-level data for {key}")
            elif d[key] is not None and abs(d[key]-value) > 0.01:
                errors.append(f"{key}: roll {value} vs independent control {d[key]}")
        if any(u["contract"] is None for u in units):
            pending.append("Scheduled contract rent is incomplete")
        unknown = [u["id"] for u in units if u["leased"] is None]
        leased = sum(u["leased"] is True for u in units)
        result["leasing"] = {"confirmed_leased": leased, "unknown_units": unknown,
                             "share": None if unknown else leased / len(units)}
        if unknown:
            pending.append("Resolve occupancy or record estimated occupancy separately")
        elif d["current_leasing"] is not None and not math.isclose(leased/len(units), d["current_leasing"], abs_tol=0.0001):
            errors.append("Current leasing disagrees with roll; resolve date, denominator or classification")
    else:
        pending.append("Unit-level rent roll unavailable")
    if deal["t12"]:
        t = deal["t12"]
        cats = {name: round(sum(vals), 2) for name, vals in t["expenses"].items()}
        amounts = [round(sum(vals), 2) for vals in zip(*t["expenses"].values())]
        # Source ledgers can allocate costs to fractions of a cent each month.
        total_expenses = round(math.fsum(v for row in t["expenses"].values() for v in row), 2)
        result["history"] = {"categories": cats, "monthly_expenses": amounts, "total_expenses": total_expenses}
        if abs(total_expenses - t["reported_expenses"]) > 0.011:
            errors.append(f"Expense detail {total_expenses} vs statement {t['reported_expenses']}")
    else:
        pending.append("Seller monthly expense history unavailable")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("deal")
    try:
        result = reconcile(read_json(parser.parse_args().deal))
        print(json.dumps(result, indent=2, allow_nan=False))
        sys.exit(int(bool(result["differences"])))
    except (ValueError, TypeError, OSError) as err:
        sys.exit(str(err))

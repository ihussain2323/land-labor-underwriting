"""Compare annual same-month changes in two Zillow ZIP CSVs from a common release."""
import argparse
import csv
from datetime import date
import json
import statistics
import sys


def observations(path, zipcode, cutoff):
    with open(path, encoding="utf-8-sig", newline="") as file:
        rows = [r for r in csv.DictReader(file) if r["RegionName"].strip().zfill(5) == zipcode]
    if len(rows) != 1:
        raise ValueError(f"{path}: expected one row for {zipcode}, found {len(rows)}")
    values = {}
    for key, v in rows[0].items():
        if v and key[:4].isdigit():
            day = date.fromisoformat(key)
            if day <= cutoff and float(v) > 0:
                values[day] = float(v)
    return values


def research(rent, value, zipcode, cutoff, baseline_year):
    series = [observations(path, zipcode, cutoff) for path in (rent, value)]
    if any(not s for s in series) or max(series[0]) != max(series[1]):
        raise ValueError("ZIP series need a common latest observation as of the cutoff")
    last = max(series[0])
    if date(baseline_year,last.month,1) <= date(2023,5,11):
        raise ValueError("Use a post-COVID baseline for this research convention")
    if baseline_year >= last.year:
        raise ValueError("At least one complete annual interval is required")
    out = {}
    for name, obs in zip(("rent", "value"), series):
        points = []
        for year in range(baseline_year, last.year + 1):
            hits = [(d, v) for d, v in obs.items() if d.year == year and d.month == last.month]
            if len(hits) != 1:
                raise ValueError(f"{name}: missing or ambiguous observation for {year}-{last.month:02}")
            day, val = hits[0]
            points.append({"date": day.isoformat(), "level": val,
                           "yoy": val/points[-1]["level"]-1 if points else None})
        out[name] = {"observations": points, "latest_yoy": points[-1]["yoy"],
                     "mean_yoy": statistics.mean(p["yoy"] for p in points[1:])}
    return out


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    for arg in ("rent_csv", "value_csv", "zipcode", "as_of"):
        p.add_argument(arg)
    p.add_argument("baseline_year", type=int)
    a = p.parse_args()
    try:
        print(json.dumps(research(a.rent_csv, a.value_csv, a.zipcode, date.fromisoformat(a.as_of), a.baseline_year), indent=2, allow_nan=False))
    except (ValueError, KeyError, OSError) as e:
        sys.exit(str(e))

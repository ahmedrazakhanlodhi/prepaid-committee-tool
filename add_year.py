#!/usr/bin/env python3
"""
add_year.py — maintainer tool to add a reporting year to the prepaid record.

The in-app upload tab was removed so viewers can't change the data. Adding a year is
now a maintainer step done here, then committed to the repo. Usage:

    python add_year.py path/to/CSPN_Prepaid_Committee_Fall_2026.xlsx 2026

It parses the committee Excel (older 2022-23 and newer 2024-25 layouts both work,
Michigan MET I/II and Mississippi tiers split automatically), merges it into
data/prepaid_master.csv (overwriting that year if already present), and updates
data/prepaid_attrs_by_year.json. Review the git diff, then commit both files.
"""
import json
import os
import sys

import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(BASE, "src"))
from prepaid_parser import parse_committee_file

DATA = os.path.join(BASE, "data")
MASTER = os.path.join(DATA, "prepaid_master.csv")
ATTRS = os.path.join(DATA, "prepaid_attrs_by_year.json")
META = os.path.join(DATA, "prepaid_meta.json")

NUM = ["funded", "assets_m", "active_accounts", "accounts_since_inception",
       "paid_out_fy_m", "paid_out_inception_m"]


def main():
    if len(sys.argv) != 3:
        print(__doc__)
        sys.exit(1)
    path, year = sys.argv[1], int(sys.argv[2])
    if not os.path.exists(path):
        sys.exit(f"File not found: {path}")

    REG = json.load(open(META))["reg"]
    recs, warns, attrs = parse_committee_file(path, year)
    for w in warns:
        print("  warning:", w)

    rows = []
    for r in recs:
        k = r["plan_key"]
        state, name, est, status, sort = REG.get(k, [k, k, "", "", 99])
        rows.append({
            "plan_key": k, "state": state, "plan_name": name, "established": est,
            "status": status, "reporting_year": year, "as_of": r.get("as_of", ""),
            "funded": r.get("funded"), "assets_m": r.get("assets_m"),
            "active_accounts": r.get("active_accounts"),
            "accounts_since_inception": r.get("accounts_since_inception"),
            "paid_out_fy_m": r.get("paid_out_fy_m"),
            "paid_out_inception_m": r.get("paid_out_inception_m"),
            "source": f"CSPN Prepaid Committee Fall {year}", "note": r.get("note", ""), "sort": sort,
        })
    new = pd.DataFrame(rows)

    master = pd.read_csv(MASTER)
    before = len(master)
    master = master[master["reporting_year"] != year]          # replace the year if present
    master = pd.concat([master, new], ignore_index=True).sort_values(["sort", "reporting_year"])
    master.to_csv(MASTER, index=False)

    by_year = json.load(open(ATTRS))
    by_year[str(year)] = attrs
    json.dump(by_year, open(ATTRS, "w"), ensure_ascii=False, indent=0)

    print(f"\nAdded {len(new)} plan rows for {year}.")
    print(f"  master: {before} -> {len(master)} rows across {master['reporting_year'].nunique()} years")
    print(f"  years now: {', '.join(str(y) for y in sorted(master['reporting_year'].unique()))}")
    print("\nReview the changes, then commit:")
    print("  git add data/prepaid_master.csv data/prepaid_attrs_by_year.json")
    print(f'  git commit -m "Add FY{year} prepaid collection"')
    print("  git push")


if __name__ == "__main__":
    main()

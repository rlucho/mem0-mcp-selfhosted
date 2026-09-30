#!/usr/bin/env python3
"""Add the standard AP tasks MA (Morocco) is missing.

Request, 2026-09-30, from the AP Team Leader for Morocco: MA's timesheet lacks
the concepts other countries have -- 5C, Others, Aging, On boarding, Xelix,
Payment.

AP is FLAT: tasks sit directly under each country project (wbs 1.x.n), with no
system level between them as Banking > Payment has. So this is one file of
top-level tasks under project 72, `idActivity` left empty.

Xelix is not created. MA already has it (#509, since before 2026-08-19). If the
team cannot see it, the cause is a missing ASSIGNMENT, not a missing activity --
the timesheet tree is built from assignments -- and creating a second Xelix
would only give them two rows to choose from.

Names are copied from the other AP countries, not from the request: AP spells
it `On boarding` (lowercase b, where Banking uses `On Boarding`), and 5C is
`DS Smith Way - 5C` in every AP country that has it.

Usage
-----
    python3 build.py --export export_Activity_YYYYMMDD_HHMMSS.csv
"""

from __future__ import annotations

import argparse
import csv
import io
import os
import sys

DELIMITER = ";"
OUTPUT_ENCODING = "cp1252"

PROJECT_ID = 72
PROJECT_NAME = "MA"
PROJECT_WBS = "1.11"

# Order follows the other AP countries (IB, UK, BE, PS ...), where these were
# created in this sequence. WBS is assigned by insertion order on import, so
# this is also the order they will appear on the timesheet.
REQUESTED = ["Aging", "Others", "Xelix", "On boarding",
             "DS Smith Way - 5C", "Payment"]

# The header the Payment kit imported successfully, values as ProjeQtor renders
# them in its own export.
COLUMNS = ["name", "idProject", "idActivity",
           "activity type", "status", "planning mode"]
ACTIVITY_TYPE = "Task"
STATUS_NEW = "recorded"
PLANNING_MODE = "as soon as possible"

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")


def read_export(path: str) -> list[dict]:
    for enc in ("cp1252", "latin-1", "utf-8-sig"):
        try:
            with open(path, encoding=enc, newline="") as fh:
                reader = csv.DictReader(fh, delimiter=DELIMITER)
                rows = list(reader)
                if reader.fieldnames and len(reader.fieldnames) > 1:
                    return rows
        except UnicodeDecodeError:
            continue
    sys.exit(f"ERROR: could not parse {path} as a ProjeQtor CSV export.")


def write_csv(path: str, rows: list[dict]) -> None:
    buf = io.StringIO(newline="")
    w = csv.DictWriter(buf, fieldnames=COLUMNS, delimiter=DELIMITER,
                       quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
    w.writeheader()
    w.writerows(rows)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(buf.getvalue().encode(OUTPUT_ENCODING))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--export", required=True, metavar="ACTIVITY_EXPORT.csv")
    args = ap.parse_args()

    rows = read_export(args.export)
    ma = [r for r in rows
          if (r.get("wbs") or "").startswith(PROJECT_WBS + ".")]
    stray = [r for r in ma if (r.get("project") or "").strip() != PROJECT_NAME]
    if stray:
        sys.exit(f"ERROR: wbs {PROJECT_WBS}.* holds rows from another project "
                 f"({stray[0]['project']!r}) -- the WBS map has moved. Refusing.")

    # Compare case- and space-insensitively: `On Boarding` already in MA would
    # still be a duplicate of `On boarding` to the person picking a timesheet row.
    norm = lambda s: " ".join(s.split()).casefold()
    have = {norm(r["name"]): r for r in ma}

    new, present = [], []
    for name in REQUESTED:
        if norm(name) in have:
            present.append(have[norm(name)])
        else:
            new.append({"name": name, "idProject": PROJECT_ID, "idActivity": "",
                        "activity type": ACTIVITY_TYPE, "status": STATUS_NEW,
                        "planning mode": PLANNING_MODE})

    print(f"MA (project {PROJECT_ID}, wbs {PROJECT_WBS}) today: {len(ma)} activities")
    for r in sorted(ma, key=lambda r: int(r["wbs"].split(".")[-1])):
        print(f"   #{r['id']:>4}  {r['wbs']:<8} {r['name']}")
    for r in present:
        print(f"already there, NOT created: {r['name']!r} (#{r['id']})")
    if not new:
        print("\nNothing to create.")
        return

    # Smoke + REMAINDER, never smoke + full: rows without an `id` always INSERT,
    # so a row present in both files would be created twice.
    smoke = os.path.join(OUT, "01a_SMOKE_create_1_task.csv")
    rest = os.path.join(OUT, "01b_REMAINDER_create_tasks.csv")
    write_csv(smoke, new[:1])
    write_csv(rest, new[1:])
    print(f"\nwrote {os.path.relpath(smoke)}   [{new[0]['name']}]  <- import FIRST")
    print(f"wrote {os.path.relpath(rest)}   "
          f"[{', '.join(r['name'] for r in new[1:])}]  <- then this")
    print(f"together: {len(new)} new tasks, no row in both files.")


if __name__ == "__main__":
    main()

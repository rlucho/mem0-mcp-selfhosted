#!/usr/bin/env python3
"""Assign the Morocco team to MA's new tasks (and to Xelix where missing).

WHO. There is no Morocco team in ProjeQtor: all 100 AP people carry team `AP`,
and nothing else in the Resource export separates the Morocco staff. So the
roster is read from the people ALREADY assigned to MA's original tasks
(#506 Posting ECM, #507 OCR IP, #508 Report IP, #509 Xelix) -- whoever works MA
today. Never the whole AP team.

WHERE. The 5 tasks build.py created, plus Xelix #509: it was on the request,
and if the team cannot see it the likely cause is exactly a missing assignment.
Gaps on #506-#508 are REPORTED, not filled -- someone missing from OCR IP may be
missing on purpose, and nobody asked for those to change.

NO DUPLICATES. A pair (task, resource) that already has an assignment in the
export -- open OR closed -- is skipped. Open means it is already there; closed
means someone removed it deliberately, and re-adding it would undo that. Rows
here carry no `id`, so each one INSERTs: the export is the only thing standing
between a re-run and a second copy of every pair. Always regenerate from a
fresh export; never re-import an old output.

The Assignment export names resources by REAL NAME, not id, so names are mapped
to ids through the Resource export, whitespace-collapsed: 55 of 214 names carry
a double space that one export may keep and the screen does not. A name that
collapses onto two ids stops the run rather than guessing.

Usage
-----
    python3 build_assignments.py --activities export_Activity_NEW.csv \\
                                 --assignments export_Assignment.csv \\
                                 --resources export_Resource.csv
"""

from __future__ import annotations

import argparse
import collections
import csv
import io
import os
import sys

import build

DELIMITER = ";"
OUTPUT_ENCODING = "cp1252"
COLUMNS = ["idProject", "refType", "refId", "idResource", "rate"]
REF_TYPE = "Activity"
RATE = 100

ORIGINAL = {"506": "Posting ECM", "507": "OCR IP", "508": "Report IP", "509": "Xelix"}
FILL_ORIGINALS = {"509"}          # Xelix: on the request. #506-#508: report only.

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "out")


def pick(row: dict, *names: str) -> str:
    """First matching column, case/space-insensitive; exits naming what it saw."""
    want = {n.replace(" ", "").lower() for n in names}
    for k in row:
        if k and k.replace(" ", "").lower() in want:
            return k
    sys.exit(f"ERROR: none of {names} in export header: {list(row)}")


def collapse(s: str) -> str:
    return " ".join((s or "").split()).casefold()


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--activities", required=True)
    ap.add_argument("--assignments", required=True)
    ap.add_argument("--resources", required=True)
    args = ap.parse_args()

    # ---- MA's tasks, from a post-import Activity export
    acts = build.read_export(args.activities)
    ma = [r for r in acts if (r.get("wbs") or "").startswith(build.PROJECT_WBS + ".")
          and (r.get("project") or "").strip() == build.PROJECT_NAME]
    by_name = collections.defaultdict(list)
    for r in ma:
        by_name[collapse(r["name"])].append(r)

    missing_orig = [f"#{i} {n}" for i, n in ORIGINAL.items()
                    if i not in {r["id"] for r in ma}]
    if missing_orig:
        sys.exit(f"ERROR: original MA tasks gone from the export: {missing_orig}")

    new_names = [n for n in build.REQUESTED
                 if collapse(n) not in {collapse(v) for v in ORIGINAL.values()}]
    new_tasks = []
    for n in new_names:
        hits = by_name.get(collapse(n), [])
        if not hits:
            sys.exit(f"ERROR: {n!r} not under MA in this export -- import "
                     f"out/01a and 01b first, then export Activity again.")
        if len(hits) > 1:
            sys.exit(f"ERROR: {n!r} exists {len(hits)} times under MA "
                     f"({', '.join('#' + h['id'] for h in hits)}) -- a create file "
                     f"was imported twice. Close the extra before assigning.")
        new_tasks.append(hits[0])
    targets = {r["id"]: r["name"] for r in new_tasks}
    targets.update({i: ORIGINAL[i] for i in FILL_ORIGINALS})

    # ---- resources: collapsed real name -> id, refusing ambiguity
    res = build.read_export(args.resources)
    rid_c, rname_c = pick(res[0], "id"), pick(res[0], "real name", "name")
    rclosed_c = next((k for k in res[0] if k and k.strip().lower() == "closed"), None)
    name_to_ids = collections.defaultdict(set)
    closed_ids = set()
    for r in res:
        name_to_ids[collapse(r[rname_c])].add(r[rid_c].strip())
        if rclosed_c and (r.get(rclosed_c) or "").strip() == "1":
            closed_ids.add(r[rid_c].strip())

    # ---- existing assignments on MA
    asg = build.read_export(args.assignments)
    if not asg:
        sys.exit("ERROR: the Assignment export is empty.")
    el_c = pick(asg[0], "element id", "refId")
    who_c = pick(asg[0], "resource", "idResource")
    type_c = next((k for k in asg[0] if k and k.replace(" ", "").lower()
                   in ("typeofitem", "element", "reftype")), None)
    cl_c = next((k for k in asg[0] if k and k.strip().lower() in ("closed", "idle")), None)

    ma_ids = {r["id"] for r in ma}
    existing = {}                 # (task id, resource id) -> "open" | "closed"
    unmapped, ambiguous = set(), set()
    for r in asg:
        el = (r.get(el_c) or "").strip()
        if el not in ma_ids:
            continue
        if type_c and (r.get(type_c) or "").strip() not in ("", "Activity"):
            continue
        raw = (r.get(who_c) or "").strip()
        ids = {raw} if raw.isdigit() else name_to_ids.get(collapse(raw), set())
        if not ids:
            unmapped.add(raw)
            continue
        if len(ids) > 1:
            ambiguous.add(f"{raw} -> {sorted(ids)}")
            continue
        state = "closed" if cl_c and (r.get(cl_c) or "").strip() == "1" else "open"
        key = (el, next(iter(ids)))
        if existing.get(key) != "open":       # any open copy wins
            existing[key] = state

    if ambiguous:
        sys.exit("ERROR: names matching more than one resource:\n  "
                 + "\n  ".join(sorted(ambiguous)))
    if unmapped:
        sys.exit("ERROR: assigned names not found in the Resource export:\n  "
                 + "\n  ".join(sorted(unmapped)))

    # ---- roster = everyone with an OPEN assignment on an original MA task
    per_task = {i: {p for (t, p), s in existing.items() if t == i and s == "open"}
                for i in ORIGINAL}
    roster = sorted(set().union(*per_task.values()) - closed_ids, key=int)
    if not roster:
        sys.exit("ERROR: nobody holds an open assignment on MA's original tasks "
                 "(#506-#509). There is no roster to copy -- the Morocco team has "
                 "to be named by hand.")

    print(f"MA roster: {len(roster)} people, from open assignments on #506-#509")
    for t, people in per_task.items():
        gap = sorted(set(roster) - people, key=int)
        fill = "  <- will fill" if t in FILL_ORIGINALS and gap else ""
        print(f"   #{t} {ORIGINAL[t]:<12} {len(people):3d} assigned"
              + (f", missing {len(gap)}: {', '.join(gap)}{fill}" if gap else ""))

    rows, skipped_closed = [], []
    for t in sorted(targets, key=int):
        for p in roster:
            s = existing.get((t, p))
            if s == "open":
                continue
            if s == "closed":
                skipped_closed.append(f"#{t} {targets[t]} / {p}")
                continue
            rows.append({"idProject": build.PROJECT_ID, "refType": REF_TYPE,
                         "refId": t, "idResource": p, "rate": RATE})

    if skipped_closed:
        print(f"\nleft alone, previously CLOSED on purpose ({len(skipped_closed)}):")
        for s in skipped_closed:
            print("   " + s)
    if not rows:
        print("\nNothing to assign -- every pair already exists.")
        return

    print(f"\n{len(rows)} new assignments:")
    for t in sorted(targets, key=int):
        n = sum(1 for r in rows if r["refId"] == t)
        if n:
            print(f"   #{t:<4} {targets[t]:<20} {n}")

    smoke = os.path.join(OUT, "02a_SMOKE_assign_1_row.csv")
    rest = os.path.join(OUT, "02b_REMAINDER_assign.csv")
    for path, part in ((smoke, rows[:1]), (rest, rows[1:])):
        buf = io.StringIO(newline="")
        w = csv.DictWriter(buf, fieldnames=COLUMNS, delimiter=DELIMITER,
                           lineterminator="\r\n")
        w.writeheader()
        w.writerows(part)
        os.makedirs(OUT, exist_ok=True)
        with open(path, "wb") as fh:
            fh.write(buf.getvalue().encode(OUTPUT_ENCODING))
    print(f"\nwrote {os.path.relpath(smoke)}      [1 row]  <- import FIRST")
    print(f"wrote {os.path.relpath(rest)}  [{len(rows) - 1} rows]")
    print("element type: Assignment. Import each ONCE -- rows without an id insert.")


if __name__ == "__main__":
    main()

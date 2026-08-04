#!/usr/bin/env python3
"""Generate the CSV import kit for the Banking > Payment (project 9) restructure.

Unlike the Collections kit next door, the top level here ALREADY EXISTS. PS, BE,
IT, DE, IB, UK, FR and PMS-TMS are sub-projects of Payment, not activities, so
nothing has to be created to hold them -- their tasks sit directly in the project
with no parent activity, which is how the 53 activities under wbs 2.1.* are
already arranged.

That splits the work in two:

  FLAT      PS, BE, IT, DE, PMS-TMS -- the target lists 11 tasks directly under
            each, and 6 of the 11 already exist. Add the missing 5. Nothing to
            close: every existing name is one of the 11, none is a stray.

  SYSTEMS   IB, UK, FR -- the target puts a system level in between (P02 + PER,
            P02 + Navision, P02 + QUALIAC), so the 11 tasks hang off a system
            rather than off the project. The system activities do not exist yet,
            and the 6 flat tasks currently sitting directly under each project
            have no place in the target shape.

Passes, because a task needs its system's numeric id and that id does not exist
until the system row has been imported:

    01_add_missing_tasks_flat.csv    ->  no export needed, parent is the project
    02_create_systems.csv            ->  export Activities  ->  03_create_tasks_under_systems.csv

Usage
-----
    python3 build.py                              # 01 and 02
    python3 build.py --export EXPORT_Activity.csv # fills in 03 as well
"""

from __future__ import annotations

import argparse
import csv
import io
import os
import sys

DELIMITER = ";"
OUTPUT_ENCODING = "cp1252"

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Banking > Payment, project id 9, wbs 2.1. Its sub-projects, from
# export_Project_20260803_142809.csv. These are PROJECTS -- an activity names its
# one in idProject and leaves idActivity empty.
PAYMENT_PROJECTS = {
    "PS": 31,
    "BE": 32,
    "IT": 33,
    "DE": 34,
    "IB": 35,
    "UK": 36,
    "FR": 37,
    "PMS-TMS": 69,
}

# Project 74, `PL & Others` (wbs 2.1.9), is also a sub-project of Payment but is
# NOT in the target structure, so nothing here touches it. Worth a decision
# separately: it holds 5 tasks, one of them named `Invoice Checks` where every
# other project uses the singular `Invoice check`.
OUT_OF_SCOPE_PROJECTS = {"PL & Others": 74}

# The 11 tasks, identical under every branch of the target structure.
TASKS = [
    "Payment Run",
    "Manual & Unplanned Payments",
    "Invoice check",
    "Other Processes",
    "Project - Robotic",
    "Mailbox",
    "Payment Run Issues",
    "Banks Issues",
    "Boarding",
    "Audit",
    "Proof of payment",
]

# Tasks hang directly off the project.
FLAT_PROJECTS = ["PS", "BE", "IT", "DE", "PMS-TMS"]

# Tasks hang off a system activity, which hangs off the project. Order matters
# only for readability; P02 is listed first everywhere in the source document.
SYSTEMS = {
    "IB": ["P02", "PER"],
    "UK": ["P02", "Navision"],
    "FR": ["P02", "QUALIAC"],
}

# Already under Payment, from export_Activity_20260803_143423.csv (wbs 2.1.*).
# Every one is a `Task` at `recorded`, directly under its project.
# Verified: each project holds exactly the first 6 of TASKS, no strays.
EXISTING = {
    "PS":      {"Payment Run": 67, "Manual & Unplanned Payments": 156,
                "Invoice check": 163, "Other Processes": 170,
                "Project - Robotic": 349, "Mailbox": 436},
    "BE":      {"Payment Run": 66, "Manual & Unplanned Payments": 157,
                "Invoice check": 164, "Other Processes": 171,
                "Project - Robotic": 350, "Mailbox": 437},
    "IT":      {"Payment Run": 65, "Manual & Unplanned Payments": 158,
                "Invoice check": 165, "Other Processes": 172,
                "Project - Robotic": 351, "Mailbox": 438},
    "DE":      {"Payment Run": 64, "Manual & Unplanned Payments": 159,
                "Invoice check": 166, "Other Processes": 173,
                "Project - Robotic": 352, "Mailbox": 439},
    "IB":      {"Payment Run": 63, "Manual & Unplanned Payments": 160,
                "Invoice check": 167, "Other Processes": 174,
                "Project - Robotic": 353, "Mailbox": 440},
    "UK":      {"Payment Run": 62, "Manual & Unplanned Payments": 161,
                "Invoice check": 168, "Other Processes": 175,
                "Project - Robotic": 354, "Mailbox": 441},
    "FR":      {"Payment Run": 10, "Manual & Unplanned Payments": 162,
                "Invoice check": 169, "Other Processes": 176,
                "Project - Robotic": 355, "Mailbox": 442},
    "PMS-TMS": {"Payment Run": 426, "Manual & Unplanned Payments": 427,
                "Invoice check": 428, "Other Processes": 429,
                "Project - Robotic": 430, "Mailbox": 443},
}

# Values as ProjeQtor renders them in its own export -- the same set the
# Collections kit imported successfully.
LABEL_ACTIVITY_TYPE = "Task"
LABEL_PLANNING_MODE = "as soon as possible"
LABEL_STATUS_NEW = "recorded"

COLUMNS = ["name", "idProject", "idActivity",
           "activity type", "status", "planning mode"]

CLOSE_COLUMNS = ["id", "name", "status"]
LABEL_STATUS_CLOSED = "closed"


def placeholder(*parts: str) -> str:
    """Marker left in idActivity until an export supplies the real system id."""
    return "<<" + " > ".join(parts) + ">>"


def row(name: str, project: str, parent="") -> dict:
    return {"name": name,
            "idProject": PAYMENT_PROJECTS[project],
            "idActivity": parent,
            "activity type": LABEL_ACTIVITY_TYPE,
            "status": LABEL_STATUS_NEW,
            "planning mode": LABEL_PLANNING_MODE}


# ---------------------------------------------------------------------------
# Row builders
# ---------------------------------------------------------------------------

def rows_flat_additions() -> list[dict]:
    """The 5 tasks per flat project that do not exist yet.

    Driven by EXISTING rather than by a count, so re-running after a partial
    import emits only what is still missing and cannot create duplicates.
    """
    out = []
    for project in FLAT_PROJECTS:
        have = EXISTING.get(project, {})
        out += [row(t, project) for t in TASKS if t not in have]
    return out


def rows_systems(ids: dict) -> list[dict]:
    """The 6 system activities, directly under their project."""
    return [row(system, project)
            for project, systems in SYSTEMS.items()
            for system in systems
            if (project, system) not in ids]


def rows_system_tasks(ids: dict) -> list[dict]:
    """11 tasks under each system. idActivity needs the system's numeric id."""
    out = []
    for project, systems in SYSTEMS.items():
        for system in systems:
            parent = ids.get((project, system), placeholder(project, system))
            out += [row(t, project, parent)
                    for t in TASKS if (project, system, t) not in ids]
    return out


def rows_close_superseded() -> list[dict]:
    """The 18 flat tasks under IB, UK and FR.

    The target gives those three a system level, so a task sitting directly under
    the project has no place in it. They are NOT deleted -- deleting would take
    their booked hours with them -- so this closes them, the same treatment the
    11 old Collections activities get.

    Held back deliberately: see README.md. Re-parenting them under P02 instead
    would keep the history in place, and that is a call to make before importing
    anything that cannot be walked back.
    """
    rows = [(aid, name) for project in SYSTEMS
            for name, aid in EXISTING[project].items()]
    # Descending id, the order the Collections kit uses: closing a parent before
    # its children errors out. These 18 are all siblings, so it costs nothing
    # here and keeps one convention across both kits.
    return [{"id": aid, "name": name, "status": LABEL_STATUS_CLOSED}
            for aid, name in sorted(rows, reverse=True)]


# ---------------------------------------------------------------------------
# Export handling
# ---------------------------------------------------------------------------

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


def resolve_ids(rows: list[dict]) -> dict:
    """Map (project, system) and (project, system, task) -> activity id.

    Scoped to wbs 2.1.* so the many same-named projects elsewhere in the instance
    (there is an `IB` under AP and another under Management) cannot be confused
    with Payment's. Depth tells the two apart: a system sits at 2.1.p.s, a task
    under it at 2.1.p.s.t.
    """
    wanted = set(SYSTEMS)
    out, seen = {}, {}
    for r in rows:
        wbs = (r.get("wbs") or "").strip()
        project = (r.get("project") or "").strip()
        name = (r.get("name") or "").strip()
        if not wbs.startswith("2.1.") or project not in wanted:
            continue
        depth = len(wbs.split("."))
        systems = SYSTEMS[project]
        if depth == 4 and name in systems:
            key = (project, name)
        elif depth == 5:
            parent_wbs = wbs.rsplit(".", 1)[0]
            system = seen.get(parent_wbs)
            if system is None or name not in TASKS:
                continue
            key = (project, system, name)
        else:
            continue
        if key in out and out[key] != r["id"]:
            sys.exit(f"ERROR: {' > '.join(key)} resolves to both #{out[key]} and "
                     f"#{r['id']}. Resolve the duplicate before importing.")
        out[key] = r["id"]
        if depth == 4:
            seen[wbs] = name
    return out


def write_csv(path: str, columns: list[str], rows: list[dict]) -> None:
    buf = io.StringIO(newline="")
    writer = csv.DictWriter(buf, fieldnames=columns, delimiter=DELIMITER,
                            quoting=csv.QUOTE_MINIMAL, lineterminator="\r\n")
    writer.writeheader()
    writer.writerows(rows)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "wb") as fh:
        fh.write(buf.getvalue().encode(OUTPUT_ENCODING, errors="replace"))


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--export", metavar="CSV",
                    help="fresh ProjeQtor activity export, used to fill in the "
                         "system ids that 03 needs")
    ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.abspath(__file__)), "out"))
    args = ap.parse_args()

    ids = resolve_ids(read_export(args.export)) if args.export else {}

    flat = rows_flat_additions()
    systems = rows_systems(ids)
    system_tasks = rows_system_tasks(ids)
    close = rows_close_superseded()

    files = [
        ("00_smoke_test_one_task.csv", COLUMNS, flat[:1]),
        ("01_add_missing_tasks_flat.csv", COLUMNS, flat),
        ("02_create_systems.csv", COLUMNS, systems),
        ("03_create_tasks_under_systems.csv", COLUMNS, system_tasks),
        ("04_close_superseded_flat.csv", CLOSE_COLUMNS, close),
    ]

    if ids:
        print(f"found in the export, so left out of the files below: "
              f"{sum(1 for k in ids if len(k) == 2)} systems, "
              f"{sum(1 for k in ids if len(k) == 3)} tasks\n")

    for filename, columns, rows in files:
        path = os.path.join(args.out, filename)
        pending = sum(1 for r in rows if str(r.get("idActivity", "")).startswith("<<"))
        # Never let a placeholder rebuild overwrite a file that already carries
        # resolved ids -- that record is the only proof of which system each task
        # was attached to. The Collections kit learned this the hard way.
        if pending and os.path.exists(path):
            with open(path, "rb") as fh:
                if b"<<" not in fh.read():
                    print(f"KEPT  {path}  [resolved ids on disk; pass --export]")
                    continue
        write_csv(path, columns, rows)
        note = f"  ({pending} parent ids still unresolved)" if pending else ""
        print(f"wrote {path}  [{len(rows)} rows]{note}")

    total = len(flat) + len(systems) + len(system_tasks)
    print(f"\nto create: {len(flat)} flat tasks + {len(systems)} systems + "
          f"{len(system_tasks)} tasks under systems = {total} activities")
    print(f"to close:  {len(close)} superseded flat tasks under "
          f"{', '.join(SYSTEMS)} (HELD -- see README.md)")
    if not ids:
        print("\nNo --export given: 03 carries <<placeholders>> and is NOT "
              "importable yet. Import 02, export Activities, then re-run "
              "with --export.")


if __name__ == "__main__":
    main()

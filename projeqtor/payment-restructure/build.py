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
    # Renamed from PS on 2026-08-06 (project record 31; see PROJECT_ALIASES).
    "NL": 31,
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

# Display names this instance has used for a project, mapped to the key the
# tables above use. Project 31 was `PS` until 2026-08-06 and is `NL` now, so an
# export taken either side of that rename still resolves. Note there is a SECOND
# `PS`, project 8 under AP (wbs 1.6), which is untouched -- every lookup here is
# confined to wbs 2.1.*, so the two can never be confused.
PROJECT_ALIASES = {"PS": "NL"}


def canon(project: str) -> str:
    """Export display name -> the key PAYMENT_PROJECTS/EXISTING are stored under."""
    return PROJECT_ALIASES.get(project, project)

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
    # Source document says "Boarding"; renamed on request before first import, so
    # no activity was ever created under the short name. Note the instance
    # already carries two other spellings elsewhere -- "On boarding" (9, on the
    # AP projects, Help Desk and Cross Tasks) and "Onboarding" (1, Human
    # Resources) -- neither of which this matches. See --rename.
    "On Boarding",
    "Audit",
    "Proof of payment",
]

# Tasks hang directly off the project.
FLAT_PROJECTS = ["NL", "BE", "IT", "DE", "PMS-TMS"]

# Earlier names that still count as "already there", so a task is not recreated
# under its new name while an export taken before the rename is in hand. Without
# this, generating between the rename file and its import emits 5 rows that would
# insert `On Boarding` alongside the `Boarding` rows about to become it.
RENAMED_FROM = {"On Boarding": ["Boarding"]}

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
    "NL":      {"Payment Run": 67, "Manual & Unplanned Payments": 156,
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

def rows_flat_additions(ids: dict) -> list[dict]:
    """The tasks per flat project that do not exist yet.

    EXISTING is only the 2026-08-03 baseline. Anything imported since is known
    only to a fresh export, so `ids` is consulted too -- without it, re-running
    re-emits rows that already went in.

    That is not hypothetical: the 1-row smoke is row 1 of this file, so importing
    the smoke and then the full file created `Payment Run Issues` twice under PS
    (#635 and #636). Rows here carry no `id`, so every one INSERTs. Pass --export
    and the file only ever holds what is genuinely missing.
    """
    out = []
    for project in FLAT_PROJECTS:
        have = set(EXISTING.get(project, {})) | {
            key[1] for key in ids if len(key) == 2 and key[0] == project}
        out += [row(t, project) for t in TASKS
                if t not in have
                and not any(o in have for o in RENAMED_FROM.get(t, ()))]
    return out


def rows_rename(rows: list[dict], mapping: list[tuple[str, str]],
                scope: str, loose: bool) -> list[dict]:
    """id;name rows renaming every activity matching `old` to `new`.

    `name` is a plain field with no workflow attached, unlike `status`, so a
    rename is a straight update -- and every row carries an `id`, so it can only
    update, never insert.
    """
    norm = (lambda s: "".join(s.split()).lower()) if loose else (lambda s: s)
    out = []
    for old, new in mapping:
        hits = [r for r in rows
                if (r.get("wbs") or "").strip().startswith(scope)
                and norm((r.get("name") or "").strip()) == norm(old)
                and (r.get("name") or "").strip() != new]
        for r in sorted(hits, key=lambda r: int(r["id"])):
            out.append({"id": r["id"], "name": new,
                        "_from": r["name"], "_project": r["project"]})
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


def flat_tasks(rows: list[dict], project: str) -> dict:
    """name -> id for tasks still sitting DIRECTLY under `project` (depth 4, not
    a system). Read from the export rather than EXISTING, because EXISTING is a
    fixed baseline that does not know the re-parenting has happened."""
    systems = SYSTEMS.get(project, [])
    out = {}
    for r in rows:
        wbs = (r.get("wbs") or "").strip()
        if (wbs.startswith("2.1.") and len(wbs.split(".")) == 4
                and canon((r.get("project") or "").strip()) == project
                and (r.get("name") or "").strip() not in systems):
            out[(r.get("name") or "").strip()] = r["id"]
    return out


def p02_duplicates(rows: list[dict]) -> list[tuple[str, str, str, str]]:
    """(new id, country, name, old id) for tasks 03 created under P02 that
    duplicate a task already sitting flat under the same project.

    03 built all 11 tasks under every system, including the 6 names that already
    existed directly under IB/UK/FR. Under the re-parent plan the OLD row wins --
    it carries the booked hours -- so its freshly created twin under P02 is the
    one that goes. The other 5 under P02 are genuinely new and stay.
    """
    p02 = {r["project"]: (r["id"], r["wbs"]) for r in rows
           if (r.get("wbs") or "").startswith("2.1.")
           and len(r["wbs"].split(".")) == 4 and r["name"] == "P02"
           and r["project"] in SYSTEMS}
    out = []
    for country, (pid, pwbs) in p02.items():
        # Compare against what is STILL FLAT, not against the EXISTING baseline.
        # Once re-parenting has run, the baseline tasks are themselves under P02,
        # so matching on the baseline makes every one of them its own duplicate --
        # and the delete list then names the rows carrying the booked hours.
        # Nothing is flat any more at that point, which is exactly right: the
        # duplicates were already dealt with.
        old = flat_tasks(rows, country)
        for r in rows:
            wbs = (r.get("wbs") or "").strip()
            if wbs.startswith(pwbs + ".") and len(wbs.split(".")) == 5:
                if r["name"] in old:
                    out.append((r["id"], country, r["name"], old[r["name"]]))
    return sorted(out, key=lambda t: int(t[0]))


def rows_reparent_to_p02(rows: list[dict]) -> list[dict]:
    """Move each old flat task under its project's P02.

    `idActivity` is a plain field -- no workflow attached, unlike `status` -- so
    this is a straight update, and every row carries an `id` so it cannot insert.
    It is also what makes closing unnecessary: the activity keeps its identity and
    its hours, and simply moves down one level.
    """
    p02 = {r["project"]: r["id"] for r in rows
           if (r.get("wbs") or "").startswith("2.1.")
           and len(r["wbs"].split(".")) == 4 and r["name"] == "P02"
           and r["project"] in SYSTEMS}
    missing = [c for c in SYSTEMS if c not in p02]
    if missing:
        sys.exit(f"ERROR: no P02 activity found under {', '.join(missing)}. "
                 f"Import 02 first, and re-export.")
    return [{"id": aid, "name": name, "idActivity": p02[country]}
            for country in SYSTEMS
            for name, aid in sorted(EXISTING[country].items(), key=lambda kv: kv[1])]


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
    out, seen = {}, {}
    for r in rows:
        wbs = (r.get("wbs") or "").strip()
        project = canon((r.get("project") or "").strip())
        name = (r.get("name") or "").strip()
        if not wbs.startswith("2.1.") or project not in PAYMENT_PROJECTS:
            continue
        depth = len(wbs.split("."))
        systems = SYSTEMS.get(project, [])
        # Depth 4 is anything sitting directly in the project: a system under
        # IB/UK/FR, a task under the flat ones. Both are keyed (project, name) --
        # no collision, since no system shares a name with a task.
        if depth == 4:
            key = (project, name)
        elif depth == 5:
            parent_wbs = wbs.rsplit(".", 1)[0]
            system = seen.get(parent_wbs)
            if system is None or system not in systems or name not in TASKS:
                continue
            key = (project, system, name)
        else:
            continue
        if key in out and out[key] != r["id"]:
            where = " > ".join(key)
            # A duplicated SYSTEM is fatal: it is a parent, and the next pass
            # would hang 11 tasks off whichever id happened to win. A duplicated
            # task is a mess to tidy but blocks nothing, so it must not stop an
            # unrelated branch from being generated.
            if len(key) == 2 and name in systems:
                sys.exit(f"ERROR: system {where} resolves to both #{out[key]} and "
                         f"#{r['id']}. Delete one before importing 03, or its "
                         f"tasks will hang off an ambiguous parent.")
            print(f"WARNING: {where} exists twice, #{out[key]} and #{r['id']}. "
                  f"Not fatal -- it is a leaf -- but delete one.")
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
    ap.add_argument("--reparent", action="store_true",
                    help="move the old flat tasks under IB/UK/FR beneath their "
                         "project's P02, keeping their booked hours on one "
                         "continuous activity, and list the now-duplicate rows "
                         "03 created under P02 for manual deletion.")
    ap.add_argument("--rename", metavar="OLD=NEW", action="append", default=[],
                    help="rename every activity called OLD to NEW. Repeatable. "
                         "Needs --export, since a rename updates by id.")
    ap.add_argument("--rename-scope", default="2.1.", metavar="WBS",
                    help="wbs prefix a rename is confined to (default 2.1. = "
                         "Payment). Pass '' to reach the whole instance.")
    ap.add_argument("--rename-loose", action="store_true",
                    help="match the old name ignoring case and spaces, so "
                         "'On boarding' and 'Onboarding' both match 'onboarding'")
    args = ap.parse_args()

    export_rows = read_export(args.export) if args.export else []
    ids = resolve_ids(export_rows) if export_rows else {}

    if args.reparent:
        if not export_rows:
            sys.exit("ERROR: --reparent needs --export; P02's id comes from it.")
        moves = rows_reparent_to_p02(export_rows)
        path = os.path.join(args.out, "04b_reparent_old_flat_to_p02.csv")
        write_csv(path, ["id", "name", "idActivity"], moves)
        print(f"wrote {path}  [{len(moves)} rows]")
        for m in moves:
            print(f"  #{m['id']:>4}  {m['name']:30s} -> parent #{m['idActivity']}")

        dups = p02_duplicates(export_rows)
        listing = os.path.join(args.out, "04c_DELETE_these_p02_duplicates.txt")
        os.makedirs(os.path.dirname(listing), exist_ok=True)
        with open(listing, "w", encoding="utf-8") as fh:
            fh.write("Delete these by hand -- the import cannot delete records.\n"
                     "Each duplicates an older task being re-parented under P02;\n"
                     "the old one carries the booked hours, these were created\n"
                     "empty by 03 and have no history to lose.\n\n")
            for new, country, name, old in dups:
                fh.write(f"#{new}  {country} > P02 > {name}   "
                         f"(duplicate of #{old})\n")
        print(f"\nwrote {listing}  [{len(dups)} to delete by hand]")
        for new, country, name, old in dups:
            print(f"  DELETE #{new:>4}  {country} > P02 > {name:30s} dup of #{old}")
        return

    if args.rename:
        if not export_rows:
            sys.exit("ERROR: --rename needs --export; it updates activities by id.")
        mapping = []
        for pair in args.rename:
            if "=" not in pair:
                sys.exit(f"ERROR: --rename wants OLD=NEW, got {pair!r}")
            old, _, new = pair.partition("=")
            mapping.append((old.strip(), new.strip()))
        hits = rows_rename(export_rows, mapping, args.rename_scope, args.rename_loose)
        if not hits:
            sys.exit("No activity matched -- nothing to rename.")
        slug = "".join(c if c.isalnum() else "-" for c in mapping[0][1].lower()).strip("-")
        # An instance-wide rename reaches well beyond Payment, so it must not
        # land on the scoped file's name and be imported by mistake.
        suffix = "" if args.rename_scope else "_INSTANCE_WIDE"
        path = os.path.join(args.out, f"05_rename_{slug}{suffix}.csv")
        write_csv(path, ["id", "name"],
                  [{"id": h["id"], "name": h["name"]} for h in hits])
        scope = args.rename_scope or "the whole instance"
        print(f"wrote {path}  [{len(hits)} rows]   scope: {scope}\n")
        for h in hits:
            print(f"  #{h['id']:>4}  {h['_project']:14s} {h['_from']!r} -> {h['name']!r}")
        return

    flat = rows_flat_additions(ids)
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
        # Depth 4 is a system under IB/UK/FR but a task under the flat projects,
        # so counting every (project, name) key as a system overstates it wildly.
        found_systems = sum(1 for k in ids
                            if len(k) == 2 and k[1] in SYSTEMS.get(k[0], []))
        print(f"found in the export, so left out of the files below: "
              f"{found_systems} systems, "
              f"{sum(1 for k in ids if len(k) == 2) - found_systems} flat tasks, "
              f"{sum(1 for k in ids if len(k) == 3)} tasks under systems\n")

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

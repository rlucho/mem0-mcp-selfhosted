# ProjeQtor V11.1.3 (DS Smith, self-hosted) — everything learned

Written to be read cold by a new session. Ordered by how much each fact costs
when it is not known.

## 1. The one rule everything else follows

**A row with an `id` UPDATEs. A row without an `id` INSERTs. Nothing dedupes.**

True for every element type. It has bitten us four times:

| What happened | Cost |
|---|---|
| A 1-row smoke test left inside its own bulk file | two `Payment Run Issues` under PS (#635/#636) |
| An assignment file imported twice | 1825 duplicate assignment records, 73 activities with 50 assignments each |
| HR macro's "Update" header had no `id` column | every update run created a second resource with the same name |
| — | re-importing any allocation file duplicates every row in it |

The corollary that matters for the HR workbook: **a Create row must carry no
`id`.** An invented id does not create anything — it updates whoever holds it.

## 2. Import mechanics

* The dropdown says *"csv file (comma separated)"*. The reader wants **`;`**.
  Encoding **cp1252**, line endings **CRLF**.
* **The `?` button on the Import Data screen prints the target object's real
  schema** — field name, type, display label. Highest-value diagnostic on the
  whole screen; use it before guessing a column name.
* Import cannot DELETE. Removal is a UI job, always.
* **WBS is recalculated on import.** Hierarchy is expressed through
  `idActivity` (numeric parent id), so building a tree takes multiple passes
  with a fresh export between each.
* Element types are separate code paths — `Activity`, `Assignment`, `Project`,
  `Resource`, `Allocation`. Never infer one's behaviour from another's.
  Demonstrated: `closed`/`idle` are ignored on Activity, writable on Assignment.

## 3. Status workflow

**Staged.** The status dropdown only lists what is reachable *from where the
item currently sits*. `closed` is two hops from `recorded`:

```
from recorded:  recorded, qualified, accepted, assigned, in progress, done, cancelled
from done:      re-opened, done, verified, delivered, validated, closed
```

So closing is `recorded → done → closed`, two import files, each needing its own
smoke row — each hop is a separate transition. `status=closed` from `recorded`
returns *"the workflow does not allow you to move this item to this status"*,
which reads like "no such status" and is not.

Closing also requires **responsible** and **result** to be set. We used
Francisco Manzanilla (id 20), the Banking manager.

`idle` on an Activity returns *"No change to update"* — accepted and silently
ignored. Confirmed by writing the opposite value to a flag demonstrably off.

## 4. Timesheets

* **Visibility is driven by assignments, not status.** No assignment → the
  activity is not on anyone's timesheet.
* One timesheet row per assignment record, so a duplicated `(activity, resource)`
  pair renders the same activity id twice. That is what "duplicate activities in
  Payments" turned out to be — the tree was intact at 211; the assignments under
  it had doubled.
* Closing an activity removes it from the tree while **keeping its recorded
  hours**. Verified on #521: closed, gone from the timesheet, 180,25 h intact.
* **Assignments with booked hours cannot be deleted** — ProjeQtor withholds the
  trash icon. Close them instead.
* Group rows still render editable cells, so never assign a group row: time
  lands on it instead of rolling up from the children.
* Untoggling auto-assignment does **not** hide an activity from the tree. Tested,
  it did not work.
* Closing an allocation auto-closes that resource's activity assignments on the
  project.

## 5. Server limits — root cause of every large-import failure

`\\dss-ib-dfs-228\xampp\htdocs\sscactconspain\`

* **`max_execution_time = 120`** caused every *"didn't send any data"* failure.
  The work **commits anyway** up to the point it died — which is why a re-run
  duplicated rather than resumed.
* **`max_input_vars = 20000`** caused the 500 on large timesheets
  (*"Input variables exceeded 20000"*).
* `php.ini` under `\\dss-ib-dfs-228\xampp\php\` had the values commented out.
* Fixed **without admin rights and without a restart** by dropping a `.htaccess`
  in the app directory — mod_php reads `php_value` per directory:

```
php_value max_input_vars 50000
php_value max_execution_time 300
php_value max_input_time 300
```

  Verified working for Silvia and Alba.

* **Do not press Start in the XAMPP Control Panel from the UNC share** — running
  `xampp-control.exe` off the share executes it on the local machine, not the
  server.

## 6. Name traps in this instance

* **55 of 214 resource names contain a double space** (`Francisco  Manzanilla`).
  HTML collapses it on screen, so a retyped name looks identical and is not.
  Match resources by **numeric id**, or by **`userName`** — never by real name.
* **11 of 214 names are destroyed by `StrConv(vbProperCase)`**:
  `Silvia De la Fuente Peña` → `De La`, `Cristina Garcia-Rojo Martorell` →
  `Garcia-rojo`.
* **Duplicate project names**: `Management` ×5, `IB` ×3, `Collections` ×2,
  `PS` ×2, `PL & Others` ×2. Verify project membership by **WBS prefix**, not by
  name — a name-based traversal reports phantom gaps.
* Ids are **not contiguous** — #636 deleted while #635 stayed; FR > PER landed
  at #812 after a gap.

## 7. What was built

**Collections** (`projeqtor/collections-restructure/`) — 109-activity tree,
2025 assignments, the 11 old flat activities closed via `done → closed` so the
recorded time survives.
*Remaining: import `04d_1_set_status_done.csv` then `04d_2_set_status_closed.csv`,
10 rows each, #521 already done.*

**Payment** (`projeqtor/payment-restructure/`) — 223 activities across 10
sub-projects, 13 system levels, every branch at exactly 11 tasks, no duplicates.
1825 duplicate assignment records identified and closed.

```python
PAYMENT_PROJECTS = {"PS":31,"BE":32,"IT":33,"DE":34,"IB":35,"UK":36,
                    "FR":37,"PMS-TMS":69,"PL & Others":74,"Marruecos":75}
SYSTEMS = {"IB":["P02","PER","E01"], "UK":["PP2","Navision"],
           "FR":["P02","QUALIAC","E01","Navision","PER"],
           "IT":["PP2","E01"], "Marruecos":["E01","PER"]}
```
`PP2` is first for IT on purpose — the first system inherits the existing 69,61 h.

**HR import** (`projeqtor/hr-user-import/`) — `NewModuleMacro.bas` (corrected
module), `ResourceLookup.bas` (resource-id lookup), `RESOURCE_IDS.md` (why
auto-incrementing ids is unsafe).
*Remaining: import both `.bas` files into the workbook in a session with Excel.*

**Paula Fernández Prieto (resource/user 163, `espaufer`), AP → R2R** —
`projeqtor/hr-user-import/out/paula_*.csv`. Resource 163 and User 163 are the
same person; one Resource file covers both. Two imports plus a manual cleanup:
`paula_01A` (element type Resource, team change) → `paula_02` (element type
Allocation, 11 R2R rows) → `paula_03` by hand in the UI (remove her 12 AP
allocations — the import cannot delete).

## 8. Generator gotchas worth not re-learning

* Resolve parent ids in **two passes** — a single pass only resolves a child
  whose parent happens to appear earlier in the export, and silently misses the
  rest (48 of 66 tasks, once).
* Drive every file off the **current export**, never a static baseline table in
  the script. Two separate bugs came from this: a delete list that named the
  *keepers*, and a close list holding 1293,30 h of real time.
* Membership checks must be **depth-independent**. Re-parenting moved 18
  activities from depth 4 to depth 5, and a `depth == 4 and id in baseline`
  check stopped matching the moment they moved.
* Emit **smoke + REMAINDER**, never smoke + full.
* Split big imports into ~365-row files. A file where every row carries an `id`
  is safe to re-run; one without ids is not.

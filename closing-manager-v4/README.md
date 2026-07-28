# Closing Manager — Analysis & V4-CIO hardened macros

Analysis of the legacy SAP/Excel VBA automation
`Closing_Manager_IP_Legacy_changed_v3.xlsm`, plus a hardened **V4-CIO** set of
VBA modules that fix the defects found during the review.

> **Nothing here executes the macro or contains any credentials or live data.**
> The VBA was extracted statically from `vbaProject.bin` with `olevba`.

---

## Contents

| Path | What it is |
|------|------------|
| **`Closing_Manager_IP_V4-CIO.xlsm`** | **The ready-to-use workbook** — the original with the V4 fixes baked in. |
| **`report/Closing-Manager-Briefing-DSSmith.pdf`** | **DS Smith-branded PDF of the decision brief** — the version to circulate. |
| `report/Closing-Manager-Briefing.html` | Same brief as a web page (published artifact). |
| `report/Closing-Manager-Briefing.md` | Same brief in Markdown. |
| `report/Closing-Manager-Briefing-DSSmith.html` | Print source for the PDF (A4, brand palette). |
| `report/make_pdf.py` | Renders the print source to PDF via headless Chromium. |
| `report/Closing-Manager-Analysis.html` | The detailed technical report (open in a browser; also published as an artifact). |
| `report/Closing-Manager-Analysis.md` | Same technical report in Markdown. |
| `report/v4-changes.diff` | Unified diff of every V4 change vs the original modules. |
| `vba-v4/` | The V4-CIO modules as source (for import, or review). |
| `vba-original/` | The original modules, extracted verbatim, for reference/diffing. |
| `build/` | The offline rebuild pipeline (Python) that produced the `.xlsm`. |

## Two ways to get V4

1. **Use the built workbook (recommended):** open `Closing_Manager_IP_V4-CIO.xlsm`.
   It is the original file with only the VBA changed — see *How the workbook was
   built* below.
2. **Import the modules yourself:** replace four modules —
   `GlobalModule` / `Admin` / `Closing` / `Printing` (steps further down). Use this
   if you prefer to apply the fix to your own current copy of the workbook. The V4
   `GlobalModule` is self-contained (the `mCloseEnv_V4` helpers are folded into it),
   so there is no separate module to import.

---

## What the macro does (one paragraph)

A Capgemini "Closing Manager" for International Paper's month-end **Record-to-Report**
close. It attaches to a running **SAP GUI** session via SAP GUI Scripting, drives
transactions to export data as flat files into the workbook's own folder, parses
them back into worksheets, posts closing entries, prints statutory reports to PDF
via **PDFCreator**, merges them with a bundled **`GiosPSMC.exe`**, and files a single
dated PDF under `C:\_Files to Transfer\MONTH END CLOSE\<year>\<mm>\`. It is a
Windows-desktop RPA screen-scraper — see the full report for the pipeline, the
dependency list, and the ranked failure points.

---

## External dependencies (Capgemini / IP-Polska servers)

The macro reaches two Capgemini-hosted systems at run time, plus some
reference-only links. Status is for this `_changed_v3` build.

**Capgemini SharePoint — SOAP web service (ACTIVELY called)**

- `https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx`
  — on-prem SharePoint `Lists.asmx`. **Live dependency:** 11 calls across `Admin`
  (7) and `GlobalModule` (4), via `spGetList/spClearList/spAddToList/spUpdateList`.
  Used by `UpdateData` and for list logging during the close (from `Postings`,
  `Printing`, `UF_CPCChange`). Windows-auth (`Environ("username")`). Needs VPN.
- `https://troom-x.capgemini.com/sites/InternationalPaper/CC/CG/_vti_bin/Lists.asmx`
  — same server, `/CC/CG/` site (Employees/SAP-ID lookup). **Commented out** in v3.

**Capgemini-Poland file server — UNC (server `pl-krabpo-fsc01`, share `ipa$`)**

- `\\pl-krabpo-fsc01\ipa$\R2R\R2R - IP EU GL West\USEFUL\pdf\merger\GiosPSMC.exe`
  — **ACTIVE, load-bearing:** copied locally on first run to provide the PDF merger.
- `\\pl-krabpo-fsc01\ipa$\R2R\R2R - IP EU\MONTH-END\CLOSING REPORTS\<year>\<mm>`
  — assigned to `FShared`, but that variable is **never read** in v3 → effectively
  dead. (`pl-krabpo` = the Capgemini Poland/Kraków site; `ipa$` = the International
  Paper share — this is the "IP Polska" server.)

**Reference only (not called by code)**

- `https://capgemini.sharepoint.com/sites/InternationalPaper/capgemini/Lists/Automations_list/Allitemsg.aspx`
  — automation-governance register, sits in each module's header comment.
- `https://capgemini-my.sharepoint.com/IP All/INTERCOMPANY RECONCILIATIONS/2017/07/NEW AR AP Intercompany Matching Master.xlsm`
  (+ its mapped-drive twin `U:\IP All\...\NEW AR AP Intercompany Matching Master.xlsm`)
  — an **external-workbook formula link**, not VBA; may prompt "update links" on open.

> *Not* dependencies: `schemas.microsoft.com/sharepoint/soap/…`,
> `schemas.xmlsoap.org/soap/envelope/`, `www.w3.org/2001/XMLSchema…` are XML
> namespace identifiers inside the SOAP envelopes — never contacted.

In V4 the three live/near-live endpoints are named constants at the top of
`GlobalModule` (`CM_SP_BASE`, `CM_MERGER_SRC`, `CM_ARCHIVE_ROOT`) — re-point them
in one place.

---

## V4-CIO — the fixes

All path/environment/endpoint settings are centralised at the top of the V4
`GlobalModule` (the `mCloseEnv_V4` helpers are folded in). The four edited modules
keep the same procedure names and call signatures, so the rest of the workbook
(buttons, other macros) is unaffected.

| # | Fix | Where | Failure it removes |
|---|-----|-------|--------------------|
| 1 | **Single working drive** — one `CM_BASE_DRIVE` used everywhere (was `D:\`-if-present in `CreatePaths` but hard-coded `C:\` in `RunClosing`). | `GlobalModule`, `Closing` | Silent print/merge failure on any PC with a `D:` drive. |
| 2 | **OneDrive / SharePoint guard** — `AssertLocalWorkbook` stops with a clear message if the workbook path is a URL. | `GlobalModule` | Cryptic mid-run failure when run from the cloud. |
| 3 | **Parent-aware folder creation** — `EnsureFolders` builds any missing ancestor folders on the correct drive. | `GlobalModule` | `CreateFolder` failing when a parent doesn't exist. |
| 4 | **PDFCreator presence check** — clear error instead of an unguarded `Shell` of the (x86) path. | `Printing` | Run-time crash when PDFCreator isn't installed. |
| 5 | **Bounded merge waits** — the two PDF wait loops get a safety cap. | `Printing` | Excel hanging forever if SAP/PDFCreator stalls. |
| 6 | **Merger fast-fail** — `CombinePDF` verifies `GiosPSMC.exe` before shelling out. | `Printing` | Silent no-output when the tool is missing. |
| 7 | **Name-clash loop fix** — `fN` now increments. | `Printing` | Infinite loop when a same-named report already exists. |
| 8 | **Preflight Check** *(new)* — one-click `PreflightCheck` reports every dependency as pass/fail before a run. | `GlobalModule` | Discovering a missing dependency halfway through a close. |
| 9 | **Endpoint constants** — the SharePoint URL, the merger UNC and the archive UNC are now named constants; the hard-coded call sites (incl. `Admin`) point at them. | `GlobalModule`, `Admin` | Hunting through code to re-point servers when a site/share moves. |

Functional close logic and the SAP command sequences are **unchanged** — the diff
is intentionally small (`report/v4-changes.diff`).

### Configuration — change locations in one place

Edit the `CONFIGURATION` block at the top of `vba-v4/GlobalModule.bas` if your
environment differs:

```vba
Public Const CM_BASE_DRIVE   As String = "C:\"                                   ' working drive for \pdf\ folders
Public Const CM_REPORT_ROOT  As String = "C:\_Files to Transfer\MONTH END CLOSE\"
Public Const CM_MERGER_SRC   As String = "\\pl-krabpo-fsc01\...\GiosPSMC.exe"    ' PDF merger (network master copy)
Public Const CM_SP_BASE      As String = "https://troom-x.capgemini.com/.../Lists.asmx"  ' SharePoint SOAP endpoint
Public Const CM_ARCHIVE_ROOT As String = "\\pl-krabpo-fsc01\ipa$\R2R\...\CLOSING REPORTS\" ' UNC archive root
```

> If your standard build genuinely stores data on `D:\`, set `CM_BASE_DRIVE = "D:\"`
> — that one line now governs *every* folder consistently. The three server
> endpoints (`CM_SP_BASE`, `CM_MERGER_SRC`, `CM_ARCHIVE_ROOT`) are likewise the
> single place to re-point Capgemini / IP-Polska infrastructure.

---

## How the workbook was built (and how it was validated)

The `.xlsm` was produced **without Excel**, by surgically rewriting only the
workbook's `xl/vbaProject.bin` (the VBA container) and copying every other part
of the file byte-for-byte. Pipeline in `build/`:

1. Decompress the original `GlobalModule` / `Admin` / `Closing` / `Printing` source
   (MS-OVBA), apply the exact V4 edits, and re-compress. The `mCloseEnv_V4`
   helpers are **folded into `GlobalModule`** (see `vba-v4/GlobalModule.bas`) so no
   new module had to be added to the compound file.
2. Rebuild each module stream keeping its original compiled-code prefix and
   appending the new compressed source.
3. Force Excel to recompile from the new source on first open: bump the
   `_VBA_PROJECT` version stamp and empty the stale `__SRP_*` compiled-cache
   streams.
4. Rebuild the compound file and drop it back into the xlsm zip.

**Validation performed here (no Excel required):**

- The compound-file writer was proven with a **byte-perfect identity
  round-trip** (rebuild the original with no edits → 116/116 streams identical).
- `olevba` re-extracts all 26 modules; exactly 4 carry the V4 code, the other 22
  are identical to the original; every V4 marker is present.
- `olefile` reads every stream; all UserForms and other modules are untouched.
- `openpyxl` opens the workbook; **only `xl/vbaProject.bin` differs** from the
  original — all sheets, forms, images, external links and the sensitivity
  label are byte-identical.

> **One caveat:** the actual *recompile-and-run in Excel* step can't be exercised
> in this Linux build (no Excel/Windows here). The technique used — invalidating
> the compiled cache so Excel rebuilds from the embedded source — is the standard
> mechanism, and everything checkable offline passes. On first open, Excel may
> take a moment to recompile, and if macro security prompts appear, choose
> **Enable Content**. If your Excel ever rejects the baked file, the module
> import route below is the guaranteed fallback and produces the identical result.

## How to install V4 (≈ 2 minutes, in Excel on Windows)

Only needed if you are **not** using the pre-built workbook. VBA source can't be
re-injected into an `.xlsm` outside Excel, so V4 ships as importable modules:

1. **Back up** the `.xlsm` first (copy the file).
2. Press **Alt + F11** to open the VBA editor. If the Project pane is hidden, press **Ctrl + R**.
3. **Replace the four edited modules.** For each of `GlobalModule`, `Admin`,
   `Closing`, `Printing`:
   - right-click the existing module → **Remove …** → **No** when asked to export;
   - right-click the project → **Import File…** → choose the matching file in `vba-v4/`.

   *(Alternatively, open each module and paste the V4 contents over the old code —
   the procedure names are identical.)* The V4 `GlobalModule` already contains the
   `mCloseEnv_V4` helpers, so there is **no separate module to import**
   (`vba-v4/mCloseEnv_V4.bas` is kept only as readable reference).
4. **Compile:** menu **Debug → Compile VBAProject**. It should compile with no errors.
5. **(Recommended)** add a button on the `START` sheet assigned to the macro
   **`PreflightCheck`**, and run it once to confirm the environment.
6. Save as macro-enabled (`.xlsm`).

### `PreflightCheck` runs standalone

Users do **not** have to start a close to check their environment. `PreflightCheck`
is a self-contained, read-only macro: it does not start the close, create folders,
write files, or drive SAP (it only detects whether a session exists). It is safe to
run at any time, as often as needed.

- **Alt + F8** → select **PreflightCheck** → **Run** (it appears in the macro list
  automatically — no setup needed), or
- assign it to a button on the `START` sheet (*Developer → Insert → Button*).

It reports pass/fail on all six preconditions in one dialog: workbook location,
working drive, SAP session, PDFCreator, the PDF merger, and the cost centre.

---

## Recommended follow-ups (not included — they need live-Excel testing)

- Wrap the SAP transaction blocks in structured error handling (log + optional
  screenshot on failure) instead of the current bare `.findById` chains.
- Bind the SAP session by **connection description** rather than `Children(0)`, so
  the macro can't drive the wrong open SAP window.
- Add `Option Explicit` to every module and resolve the resulting undeclared-variable
  warnings. Decide what to do with `FShared` (now built from `CM_ARCHIVE_ROOT` but
  still never read downstream — either wire up the archive-copy step or delete it).

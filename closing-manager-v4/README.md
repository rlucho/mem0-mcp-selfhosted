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
| **`report/Closing-Manager-Executive-Brief-DSSmith.pdf`** | **2-page DS Smith-branded executive brief** — the short version to circulate. |
| `report/Closing-Manager-Briefing-DSSmith.pdf` | Full 6-page DS Smith-branded briefing (same findings, more detail). |
| `report/Closing-Manager-Briefing.html` | Same brief as a web page (published artifact). |
| `report/Closing-Manager-Briefing.md` | Same brief in Markdown. |
| `report/Closing-Manager-*-DSSmith.html` | Print sources for the two PDFs (A4, brand palette). |
| `report/make_pdf.py`, `report/make_pdf_exec.py` | Render the print sources to PDF via headless Chromium. |
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
| 8 | **Preflight Check** *(new)* — one-click `PreflightCheck` reports every dependency as pass/fail before a run, on its own `Preflight` sheet with a button. | `GlobalModule` | Discovering a missing dependency halfway through a close. |
| 9 | **Endpoint constants** — the SharePoint URL, the merger UNC and the archive UNC are now named constants; the hard-coded call sites (incl. `Admin`) point at them. | `GlobalModule`, `Admin` | Hunting through code to re-point servers when a site/share moves. |
| 10 | **SAP authorisation sweep** *(new)* — optional part of `PreflightCheck`: tests the 8 transactions, 6 SE16 tables, 4 GR55 report groups and 4 ALV display layouts, plus an optional read-only company-code data probe. | `GlobalModule` | Hitting an authorisation wall — or a missing layout — mid-close, after entries are posted. |
| 11 | **Bounded print waits** — the ten “wait for the printed PDF” loops now use `CM_WaitForPrint`: `Dir`-based, capped at 240 s + 180 s, `DoEvents` every second. | `Printing` | **The freeze.** The old loop rebuilt a `Shell.Application` on every pass *with no pause*, so an empty temp folder became a tight spin — Excel “Not Responding”, and eventually `-2147417848 (80010108) Method 'NameSpace' of object 'IShellDispatch6' failed`. |
| 12 | **Live progress** — `CM_Begin` / `CM_Note` / `CM_Done` report the current stage (`[14/27] printing ZGE1174 … running 06:12 — press Esc to stop`) on Excel's status bar; every wait ticks a counter. | `GlobalModule`, `Closing` | “Is it working or has it hung?” — previously indistinguishable. |
| 13 | **Plain-language failure dialog** — `CM_Explain` classifies the failure as DATA / FILE / SAP / PRINTING / EXCEL / TECHNICAL and shows *what it was doing, what went wrong, why it usually happens, what to try*, plus a technical line for the CI Team. Armed once via `On Error GoTo CM_Fail` at the top of `RunClosing`. | `GlobalModule`, `Closing` | Bare `Run-time error 13` dialogs with no context. |
| 14 | **Locale-proof amounts — all 10 sites** — `CM_AmountReq` (9 text-parsed sites) / `CM_Amount` (the ZGLRME worksheet column). Reads the literal, not the locale: an already-numeric value passes through untouched; with both separators the last one is the decimal point; a repeated single separator is grouping (`1.234.567` → 1234567, either convention); anything else is the decimal point. | `GlobalModule`, `Closing`, `Printing` | `Run-time error 13: Type mismatch` when SAP's decimal notation ≠ the PC's Windows regional format — and, worse, `1.234,56` being read silently as `1.23456`. `Round("1.234,56", 2)` and `CDbl("1.234,56")` are both locale-dependent, so this bit in a different report each time. |
| 20 | **The failure names the line it failed on** — every extract the close reads is announced with `CM_Source`, every line with `CM_Reading`, so an amount failure reports the **file, the line number, and the SAP line verbatim** — which carries the document number, account and profit centre. A worksheet-sourced failure names the sheet and row instead. The whole message is appended to `ClosingManager_errors.log` beside the workbook. | `GlobalModule`, `Closing`, `Printing`, `Postings` | "A value could not be read" with no way to find which one. |
| 19 | **The decimal convention resolves itself** — `CM_ToAmount` learns which character SAP is using as the decimal point from the first unambiguous amount in a run (anything with both separators, a repeated separator, or 1/2/4+ decimals) and applies it to any ambiguous one afterwards. `CM_SAP_DECIMAL` (default `"auto"`) pins it permanently if a site ever needs to. | `GlobalModule` | Having to stop and ask a human to resolve `1,234` — the macro now works it out itself, and only stops if nothing in the run says which convention is in force. |
| 17 | **`Postings.bas` brought into V4** — it was never rebuilt before. It carried **five more** of the freeze loops (fix 11) — one inside `Check_ZGE132AP`, which runs *immediately after* the entries go into SAP — and **eleven more** locale-dependent amount conversions (fix 14). Both now fixed there too. | `Postings` | A freeze or a `Type mismatch` on the post-and-verify path, where the entries are already real but the run never reaches the check that would say so. |
| 18 | **`Post_ZGLGWUL` truncation bug** — the negative branch read `Left(arr(U), Len(arr(U - 1)))`: the length of the **previous** column, so a negative account-44400200 amount was cut to the wrong number of characters whenever the two columns differed in width. | `Postings` | A silently wrong figure in the ZGLGWUL posting. |
| 16 | **`/arek2` downgraded to informational** — a `~` prefix in `CM_SAP_LAYOUTS` marks a layout as attempted-but-never-failing. `ZGR215 /arek2` is one: the close reaches that field only after the document-number popup, so it cannot honestly be tested from a cold selection screen. | `GlobalModule` | A false `[X]` sending people to look for a problem that was not there. |
| 15 | **Print + merge rehearsal** *(new)* — optional part of `PreflightCheck`: prints two test pages (SE16/T001 display via SAP, or Excel as fallback), confirms PDFCreator auto-saves into `\pdf\temp`, then merges them with `GiosPSMC.exe`. Also checks the Windows printer is actually named `PDFCreator`, and offers to clear leftovers from `\pdf\temp`. **The merge can also be tested on its own** (answer *No* at the prompt): `CM_SeedPdf` writes two built-in one-page PDFs straight to disk, so the merger is proved with no printer involved and a broken PDFCreator cannot mask a broken merger. | `GlobalModule` | Discovering *during* a close that PDFCreator isn't auto-saving, or that the merger is broken — the exact conditions behind fix 11. |

Functional close logic and the SAP command sequences are **unchanged** — the diff
is intentionally small (`report/v4-changes.diff`).

> **Nothing is written or posted to SAP by any preflight check.** The
> authorisation sweep navigates screens; the data probe runs `ZGLRME` with the
> transmit flag cleared; the print rehearsal displays `SE16`/`T001` and prints it.
> Printing does raise a temporary SAP print job (deleted after printing) — that
> is the only trace it leaves, and `PreflightCheck` asks before doing it.

### Number formats: how the macro is agnostic

Nothing in the close reads a number through Windows any more. Amounts are parsed
from the literal, in this order:

| The value | Read as | Why |
|---|---|---|
| already numeric | itself | a worksheet cell needs no parsing |
| `1.234,56` / `1,234.56` | 1234.56 | both separators present — the **last** one is the decimal point |
| `1.234.567` / `1,234,567` | 1234567 | a separator repeated must be grouping; a number has one decimal point |
| `1234,56` / `1234.56` / `1234,5678` | as written | one separator, 1/2/4+ digits after → decimal point |
| `1,234` / `1.234` | **learned** | genuinely ambiguous — see below |
| `1 234,56` | 1234.56 | spaces and hard spaces are grouping |
| `1234,56-` / `-1234,56` | −1234.56 | trailing or leading minus |
| `(1234.56)` | −1234.56 | the report-painter convention |

**The ambiguous case resolves itself.** `1,234` is 1234 to an English reader and
1.234 to a German one, and the string alone cannot say which. So `CM_ToAmount`
records the convention every time it sees an *unambiguous* value and applies it
to the ambiguous ones:

```
1.234,56    -> 1234.56        convention learned: ","
987,00      -> 987.00
12.000      -> 12000
1.234       -> 1234           <- resolved by what was learned above
45.678.901  -> 45678901
```

Learning is reset at the start of each run (`CM_Begin`). The only case that still
stops is an ambiguous value arriving **before** anything unambiguous has been seen
— and the dialog then names the one-line cure: set `CM_SAP_DECIMAL` to `"."` or
`","` and it never has to work it out again. `PreflightCheck` reports which
convention is in force.

The macro will not guess. A figure wrong by a factor of 1000 in a signed report
pack is worse than a stop.

### Finding the figure that failed

The dialog no longer stops at "a value could not be read". Every SAP extract is
announced as it is opened and every line as it is read, so a failure can say
exactly where to look:

```
WHAT IT WAS DOING
        step 20 of 27 - running ZGLGWUL

WHAT WENT WRONG   (DATA)
        SAP sent an amount that could mean two things, and
        nothing else in this run said which.
        The value was:  [1,234]
        That is either 1234 or about 1.

WHERE TO LOOK
        Extract file:   zglgwul.txt   -   line 147
        This is the line SAP sent (it carries the document
        number, account and profit centre):
        |  44400200 |  1234567890 | PC4711  | ...  |    1,234 |

WHAT TO TRY
        This one is settled permanently by a single setting.
        ...

A copy of this message was saved to:
C:\Closing\ClosingManager_errors.log
```

- **Text extracts** (`zglrme.txt`, `zge132.txt`, `zge132G.txt`, `zglgwul.txt`,
  `gtb1.txt`, `eis4.txt`, `gis4.txt`, `aa02.txt`, `zge1174.txt`, `zge132gwul.txt`)
  report file + line number + the line verbatim. The line is captured *before* the
  macro collapses its spacing, so the columns still line up and the document
  number is readable.
- **Worksheet-sourced values** (the ZGLRME amount column) report the sheet and the
  real sheet row — open the `ZGLRME` tab and go straight to it.
- **A print that never produced a PDF** reports the report name, the folder being
  watched, and how long it waited.
- Everything is **appended** to `ClosingManager_errors.log` next to the workbook,
  with a timestamp and the user name, so a sequence of failures across one close
  can be sent to the CI Team in one go rather than retyped from a dialog someone
  already clicked away.

### What "Balance Control Entry not completed…" means

This is the **original macro's own check**, unchanged in V4 — one of the ten
`UF_Error` dialogs the audit above confirms are still there. It is working
correctly when it fires.

**What it detects.** After the macro posts its balancing entries, `Check_ZGE132AP`
re-runs `ZGE132` in SAP and re-reads the balances. If the profit-centre amounts
still do not net to zero, `CPCL` (local currency) or `CPCG` (group currency) is
set, the offending rows are written to the **`Errors` sheet**, and the run stops.
The only way the numbers move between the macro's read and its re-read is
**something else posting into the company code while the close is running** —
which is exactly what the message says.

**What the macro does about it.** Shows the dialog and `Exit Sub`. Nothing further
is printed, nothing is merged, no report pack is filed. Entries already posted stay
in SAP — they are real documents.

**Should the whole thing be run again? Yes.** The macro does *not* leave a resume
marker for this error. The mid-run resume path exists only for the batch-input
"XY" failure, which writes `XY-<printN>` into `config!AA12` / `AA17`; on the next
run `RunAgainXY` picks that up and skips the printing already done. This error
writes `CPCL` / `CPCG`, not `XY-…`, so the next run starts from the top and
reprints everything — which is what you want, since the partial pack is not a
complete set.

**Re-running does not double-post.** `Post_ZGE132` does not replay a stored amount.
It drives `ZGE132` in SAP with `P_CLENT` and `P_BCENT` ticked and lets SAP compute
the entries from the **live** balances at that moment. So the second run posts only
whatever is still unbalanced — and nothing at all if the interfering document has
been reversed and everything nets.

**Before re-running:** make sure the interfering posting is genuinely dealt with,
and that nobody posts into that company code while the close runs. Otherwise the
same check fires again — correctly.

### Safety: nothing that used to stop the close was removed

Every construct that abandons a run was matched between the original and V4 by its
**guard condition**, not its line number:

| | original | V4 | |
|---|---|---|---|
| `Exit Sub` / `Exit Function` / `End` | 20 | 56 | all originals still present; the extras are new guards |
| `MsgBox` warnings | 0 | 12 | all new |
| `UF_Error` dialogs (the macro's own) | 10 | 10 | unchanged |
| Balance / difference checks (`ZGLRMEErr`, `Round(Bal…)`, `CCOpened`, `CheckZGLRME`, `CheckAA02`, `CheckZGE132`) | — | — | unchanged, one for one |

Two `Exit Sub` sites read as "missing" on a naive text diff because single-line
`If … Then Exit Sub` became a four-line block so `CM_Done` could clear the status
bar. The conditions are byte-identical:

```vba
If SAPID = "" Then Exit Sub                                  ' original
If SAPID = "" Then                                           ' V4
    CM_Done
    Exit Sub
End If
```

The ten `If size1 = size2 And size1 <> 0 And size2 <> 0 Then` checks also read as
removed. They were not: `CM_WaitForPrint` carries the same test
(`If s1 >= 0 And s1 = s2 And s1 > 0 Then Exit Do`) plus an existence check the
original never had.

#### The one place V4 deliberately stops where the original did not

An **ambiguous** amount. `1,234` is 1234 to an English reader and 1.234 to a German
one, and nothing in the string says which — the original guessed from Windows, which
is exactly how a figure ends up wrong by a factor of 1000 in a signed report pack.
V4 stops and explains instead. This cannot fire on a 2-decimal currency (the last
separator is followed by two digits, not three), which is every currency this close
handles.

And the mirror of that: a **blank** amount parsed out of a SAP text export used to
raise `Type mismatch` and stop the close. It still stops — `CM_AmountReq` is the
strict variant used at all nine text-parsed sites, precisely so a missing amount
cannot be silently read as zero and let an out-of-balance close pass. Only the
ZGLRME worksheet column uses the lenient `CM_Amount`, matching the original, where
an empty cell means a row with no amount.

`build/check_amount.py` is the harness: it runs 23 value shapes through a model of
the original (under both en-US and de-DE Windows) and through V4, and flags any case
where V4 continues but the original stopped, or returns a different number.

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

### `PreflightCheck` — dedicated sheet, button, and standalone run

The built workbook has a **`Preflight` sheet** (second tab) with a **Run Preflight
Check** button wired to the macro. The macro name is unchanged — the button simply
calls `[0]!PreflightCheck`, the same convention the workbook already uses for
`[0]!UpdateData` on the START sheet.

Users can also run it without the sheet: **Alt + F8** → **PreflightCheck** → **Run**.

**Part 1 — environment (instant, read-only).** Does not start the close, create
folders or write files. Reports pass/fail on: workbook location, working drive,
SAP session, PDFCreator (both the program **and** a Windows printer whose name
contains `PDFCreator` — SAP is told that exact name), the PDF merger, and the
company code.

**Part 2 — SAP authorisations (optional, asks first).** If a SAP session is
detected, the check offers to test everything the close drives in SAP:

| Checked | Objects |
|---|---|
| Transactions | `SE16` `GR55` `SM35` `ZGE132` `ZGLRME` `ZGR215` `ZGLGWUL` `ZGE1174` |
| Tables (via SE16) | `T001` `T001B` `T001Z` `SKB1` `ZCCOD` `ZGXMIT` |
| GR55 report groups | `AA02` `EIS4` `GIS4` `GTB1` |
| Report layouts (the `/…` variants) | `ZGLRME` → `/default` (`P_VARID`), `/closing` + `/default` (`P_VARIE`); `ZGR215` → `/arek2` (`P_ALV`, informational) |

It navigates to each object and reads the status bar, reporting any error message
**verbatim** rather than matching keywords — so it works whatever the SAP logon
language is. It runs no report and posts nothing, but it *does* move the SAP session
between screens (and returns it to `/n` afterwards), which is why it asks first
rather than running automatically with the environment checks.

Keep the three lists in `GlobalModule` (`CM_SAP_TCODES`, `CM_SAP_TABLES`,
`CM_SAP_RGROUPS`) in step with the code if a transaction, table or report group is
ever added.

**Layouts are checked too.** The `/default`, `/closing` and `/arek2` values are
**ALV display layouts** (typed into `P_VARID` / `P_VARIE` / `P_ALV`), not
selection-screen variants. For each one the sweep opens the transaction, **fills the
mandatory selection fields first** (`CM_FillContext` — company code and period, with
`P_XMIT` cleared so nothing is transmitted), confirms the layout field still exists,
enters the layout and presses **Enter only**. The report is never executed, so
nothing is selected, written or posted.

> Earlier builds skipped the context fields and pressed Enter on an empty selection
> screen. ZGLRME answered *"Enter Transm. Group or Comp. Code or Profit Center or
> Cost Center"* — a mandatory-field message, nothing to do with the layout — which
> was reported as three false `[X]` failures. Fixed.

> **`ZGR215 /arek2` is reported `[~]`, never `[X]`.** The close only reaches that
> layout field *after* ZGR215's document-number popup, so it cannot be tested from
> a cold selection screen and a negative answer there means nothing. A `~` prefix
> on the entry in `CM_SAP_LAYOUTS` marks it informational — still attempted, never
> counted as a failure. Use the same prefix for any future layout in the same
> position.

**Company-code data access is checked too — behind a second, separate consent.**
Reaching a transaction is not the same as being allowed to read a given company
code's data, and that only surfaces once a report actually selects. So the sweep
offers a final probe: it runs **ZGLRME** for the configured company code and
period with the transmit flag (`P_XMIT`) **cleared** — it selects and displays,
and posts nothing.

Two transactions are deliberately **never** executed, because they write:

| Excluded | Why |
|---|---|
| `SM35` | processes batch-input sessions — that is how the close posts |
| `ZGLGWUL` | carries a `P_POST` checkbox the macro ticks to post |

> **Note on `config!B2`:** it is the **Company Code**, not a cost centre — it is
> typed into `S_BUKRS-LOW`, `P_BUKRS`, `SBUKRS`, `P_CCODE` and `_COCODES-LOW`,
> which are all company-code fields. The preflight labels it correctly.

> A selection screen may answer a bare Enter with "fill in required fields";
> because SAP's message is passed through verbatim you can see that is what
> happened rather than mistaking it for a missing layout.

**Part 3 — print + merge rehearsal (optional, asks first).** Every other check is
a *look*; this one is a *rehearsal*, because the only way to know the printing
chain works is to run it:

```
SAP  →  front-end printer LOCLX  →  Windows printer "PDFCreator"
     →  auto-saved PDF in C:\pdf\temp  →  GiosPSMC.exe merge
```

The prompt offers three answers:

| Answer | What runs |
|---|---|
| **Yes** | print two test pages → wait for each PDF → merge them. Proves the whole chain. ~2 min. |
| **No** | **merge only.** Two built-in one-page PDFs (`CM_SeedPdf`) are written straight to disk and merged. No printer, no SAP, a few seconds. |
| **Cancel** | skip. |

Either way it reports each stage and the resulting file sizes, checks the merged
file really is a PDF (`%PDF-` … `%%EOF`), and deletes everything it wrote under
`C:\pdf\`. If the print stage fails, the merge is **still** tested using the
built-in pages — a broken PDFCreator cannot hide a broken merger, or vice versa.

- **Source of the test pages:** `SE16` → table `T001` (company-code names),
  limited to 3 rows and printed. A *display*, so no business data is created or
  changed; the print does raise a temporary SAP print job. If SAP is unavailable —
  or its print dialog doesn't appear — it falls back to printing a page from Excel
  and says which route it used. Either route still proves PDFCreator and the merger.
- **Leftovers are a finding.** The close treats whatever is sitting in
  `C:\pdf\temp` as its freshly printed report, so leftovers can end up inside the
  report pack. If any are found the check reports them and offers to clear them.

This is the check that catches the condition behind the freeze: **PDFCreator
installed but not auto-saving into `C:\pdf\temp`** (or waiting on a dialog). The
macro's `SetPDFCreator` only *launches* PDFCreator — every registry write that
would configure the auto-save folder is commented out in the original code, so
that configuration is an unverified assumption on every PC. Now it is verified.

### What the user sees while it runs

`RunClosing` reports its stage on Excel's status bar and refreshes it every second
during any wait:

```
CLOSING MANAGER   [14/27]   printing ZGE1174   -   waiting for the PDF of ZGE1174 (37s)      (running 06:12 - press Esc to stop)
```

### What the user sees when it stops

Instead of a bare `Run-time error 13`, `CM_Explain` shows:

```
CLOSING MANAGER - COULD NOT CONTINUE
--------------------------------------------

WHAT IT WAS DOING
        step 9 of 27 - checking ZGLRME and AA02 for differences

WHAT WENT WRONG   (DATA)
        SAP sent an amount the macro could not read as a number.
        The value was:  1.234.567,89
        On data row:    412

WHY THIS USUALLY HAPPENS
        Almost always the number format. SAP writes amounts using the
        SAP user's decimal notation, and this PC reads them using its
        Windows regional settings. ...

WHAT TO TRY
        1. SAP: System > User Profile > Own Data > Defaults > Decimal Notation ...
        2. Windows: Settings > Time & language > Region > Regional format ...
        3. If the value above is not a number at all (for example *****),
           the SAP report column is too narrow - tell the CI Team.

--------------------------------------------
This failure itself posted nothing. Anything already posted to SAP
earlier in this run stays posted - check before running again.

For the CI Team:  error 13 - Type mismatch
```

Failures are classified as **DATA** (number formats, empty selections),
**FILE** (working files, OneDrive, leftovers), **SAP** (screen not as expected,
authorisations), **PRINTING** (no PDF appeared), **EXCEL** (sheet renamed or
protected) or **TECHNICAL** (send it to the CI Team) — so a user can tell at a
glance whether it is something they can fix.

---

## Recommended follow-ups (not included — they need live-Excel testing)

- Wrap the SAP transaction blocks in structured error handling (log + optional
  screenshot on failure) instead of the current bare `.findById` chains.
- Bind the SAP session by **connection description** rather than `Children(0)`, so
  the macro can't drive the wrong open SAP window.
- Add `Option Explicit` to every module and resolve the resulting undeclared-variable
  warnings. Decide what to do with `FShared` (now built from `CM_ARCHIVE_ROOT` but
  still never read downstream — either wire up the archive-copy step or delete it).

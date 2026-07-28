# Can we run the month-end Closing Manager on our own?

*A plain-language look at what this Excel macro does, what it depends on, and the
two outside systems we should cut loose to make it fully ours.*

---

## The short answer

**Almost — and the gap is smaller than it looks.** The macro's real work (driving
SAP, printing and merging the reports) already happens on the user's own PC. But it
still reaches out to **two Capgemini-hosted systems** for settings, logging and one
small program. Neither holds anything we couldn't own ourselves, so both can be
brought in-house.

| | |
|---|---|
| **2** | outside systems it still needs — both Capgemini-hosted |
| **4** | settings tables living on their SharePoint, not in our file |
| **0** | of those dependencies are technically hard to replace |

---

## 1. What it does

An Excel VBA macro that runs the month-end close end to end:

- **Drives SAP** via SAP GUI Scripting — runs the closing transactions (SE16 table
  extracts, `ZGLRME`, `ZGE132`, `AA02`, `ZGLGWUL` and others), exports the data to flat
  files, reads it back into the workbook and posts the closing entries.
- **Creates the PDFs** — sets SAP's front-end printer to `PDFCreator` and prints each
  report out to `\pdf\final\`.
- **Merges the PDFs** — calls `GiosPSMC.exe` to concatenate the report set into one
  pack, then files it as `<CC> Closure <MM> <YYYY>.pdf` under
  `C:\_Files to Transfer\MONTH END CLOSE\`.
- **Reads and writes configuration** over SOAP to a Capgemini SharePoint site — the
  subject of this review.

One design point drives most of the fragility: SAP is driven **by screen element ID**,
not through a BAPI or other interface. The macro is a screen-scraper, so it is
sensitive to SAP UI changes, unexpected pop-ups and authorisation differences.

## 2. What it depends on

**🔴 Capgemini SharePoint — NOT OURS**
A Capgemini intranet site holding the macro's **settings tables**, which also
receives a **log of every close we run**.
`https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx`
*Why it matters:* if the site is retired, or we come off their network, the macro
quietly loads **no** settings — it does not show an error. It also means our closing
activity is recorded on a vendor's system.

**🔴 Capgemini Poland file server — NOT OURS**
A network drive at Capgemini's Kraków site. The macro copies one small program from
it — `GiosPSMC.exe`, the tool that merges the PDFs into a single report.
`\\pl-krabpo-fsc01\ipa$\R2R\R2R - IP EU GL West\USEFUL\pdf\merger\GiosPSMC.exe`
*Why it matters:* this is only a file copy, and only the first time. Once the tool is
on the PC, the server is never needed again — so this one is easy to remove for good.

**🟢 SAP — ours / business-critical.** Must be open, logged in, with scripting on.
Legitimate dependency; it stays.

**🟠 PDFCreator on the PC — ours to install.** Free software; the printer must be
named exactly `PDFCreator`.

**🟢 Local folders — ours.** `C:\pdf\` for working files and
`C:\_Files to Transfer\MONTH END CLOSE\` for the finished report. Note the workbook
must sit in a **normal local folder** — running it from OneDrive/SharePoint breaks it.

## 3. What is actually stored on their SharePoint

This is the crux of the independence question. It is not accounting data — it is
**configuration and an activity log**:

| Table | What it holds | How it's used |
|---|---|---|
| `CCCrossList` | Per cost centre: posting block, location-closed flag, which postings are allowed | Read at startup; also edited from a form inside the workbook |
| `ClosingVariants` | SAP variant name and currency per cost centre | Read at startup |
| `ProfitCenters` | Profit-centre reference data | Read at startup; refreshed by the macro |
| `ClosingTracker` | Who ran a close, for which cost centre and period, when, success/failure | Written during runs; checked to confirm data was refreshed |

All four are **small lists of settings** — the kind of thing that fits comfortably in
a worksheet or a list on our own SharePoint. Nothing here requires Capgemini's
infrastructure.

### ⚠️ Two things worth raising with Capgemini

1. **Our close activity is logged to their system.** Every run writes user name, cost
   centre, period and timestamp to `ClosingTracker` on their SharePoint. We should
   confirm this is intended and agreed, and who can see it.
2. **Failure is silent.** The macro never checks whether the SharePoint call
   succeeded. If the site is gone or we're off the network, it loads no settings and
   carries on — rather than stopping and telling the user. Fix this first, regardless
   of the ownership question.

> **Encouraging sign:** whoever produced this "v3" version was already heading this
> way. Several SharePoint calls are switched off in the code — including the lookup
> that used to fetch the user's SAP ID, now replaced by simply asking the user. The
> decoupling has started; it just wasn't finished.

## 4. What breaks it today

| Problem | What the user sees | Status |
|---|---|---|
| **The "D: drive" bug** (a real coding error) | On a PC with a D: drive, folders are created on C: but read from D:. The close appears to work until printing — which silently produces nothing. | Fixed in V4 |
| **Running from OneDrive/SharePoint** | The macro saves temp files next to itself; from a cloud location its "folder" is a web address, so the run dies early. | Guarded in V4 |
| **SAP not open / scripting off** | Clear message: "You are not logged in SAP." | Handled |
| **SAP screens change** | The macro clicks SAP by exact screen positions; an upgrade or pop-up stops it mid-run. | Ongoing risk |
| **PDFCreator missing/renamed** | Previously crashed with a technical error; now reports clearly. | Improved in V4 |
| **Off the network / no VPN** | Settings silently don't load; merge tool can't be fetched if not already local. | **Needs decision** |

## 5. Preflight Check — can users run it on its own?

**Yes. It runs standalone, and it changes nothing.**

`PreflightCheck` is a self-contained macro added in V4 specifically so users can verify
their environment **before** committing to a close. It does not start the close, does
not touch SAP beyond detecting whether a session exists, creates no folders, and writes
no files. It is safe to run at any time, as often as needed.

**How to run it:** press **Alt + F8**, select **PreflightCheck**, click **Run**. It
appears in the macro list automatically. For everyday use, add a button on the `START`
sheet (*Developer → Insert → Button*) and assign it to the same macro.

It reports on all six preconditions in one dialog:

```
CLOSING MANAGER  -  PREFLIGHT CHECK
--------------------------------------
[OK] Workbook location is local.
[OK] Working drive C:\ is present.
[X]  SAP GUI not logged in, or GUI scripting is disabled.
[OK] PDFCreator is installed.
[~]  PDF merger will be copied from the network share on first run.
[OK] Cost Centre = 1234.
--------------------------------------
NOT READY - resolve the [X] items above.
```

This turns every dependency in section 2 into a plain pass/fail before anyone starts a
period close — including whether the network-hosted merge utility is still needed on
that machine.

## 6. How we become fully self-sufficient

Each step stands on its own, so they can be done one at a time.

1. **Take a permanent copy of the PDF merge tool.** Copy `GiosPSMC.exe` off their
   share and keep it somewhere we own, or deploy it with the workbook. Quickest win.
   Confirm with Capgemini what the tool is and how it's licensed before redistributing.
2. **Ask Capgemini to export the four settings tables** (`CCCrossList`,
   `ClosingVariants`, `ProfitCenters`, `ClosingTracker`), plus who maintains them
   today and how often they change.
3. **Decide where those settings should live** — hidden worksheets inside the
   workbook (simplest, no server at all), or a list on *our* SharePoint. Depends on
   how many people edit them and whether we want an audit trail.
4. **Repoint or switch off the SharePoint calls.** V4 already puts the address in a
   **single named setting** (`CM_SP_BASE`), so re-pointing is a one-line change.
5. **Make failure loud instead of silent.** Check the settings actually loaded, so a
   dead endpoint stops the run rather than closing the books on empty configuration.
   The most important safety fix on this list.
6. **Clarify ownership and support of the macro itself.** The code names an owner and
   reviewer on a Capgemini governance register. If this is becoming ours, we need the
   source, the right to change it, and a named supporter.

> **Already delivered:** a hardened **V4-CIO** workbook fixes the D:-drive bug, blocks
> running from OneDrive, stops the macro hanging forever if a step stalls, and adds a
> one-click **Preflight Check** that tells the user whether SAP, PDFCreator, the
> folders and the network are all in place *before* they start a close. It also
> gathers the external addresses into named settings, so the work above becomes a
> configuration change rather than a code rewrite.

### Where this leaves us

Nothing in this automation genuinely requires Capgemini's servers. What lives there is
**configuration we should own anyway** and **a log of our own activity**. The
realistic path: take the merge tool, take the settings, point the macro at our own
home for them, and make it complain loudly if anything is missing. After that, the
only outside system left is SAP — exactly as it should be.

---|---|
| Macro / VBA | Program code stored inside the Excel file itself; runs when someone opens the workbook and clicks a button. |
| SAP GUI Scripting | An SAP feature letting a program press buttons in the SAP window automatically. Must be enabled on the PC and by SAP admins. |
| SharePoint list | A simple table on an intranet site — like a shared spreadsheet programs can read and write. |
| UNC path | An address for a shared network folder, `\\server\folder`. Only reachable on the company network or VPN. |
| Cost centre / close | The month-end routine of finalising the books for part of the business; the macro produces the report pack evidencing it. |
| Preflight Check | The new V4 button that verifies everything the macro needs is ready before a run starts. |

---

*Based on a full read of the macro's source code. No macro was run and no live system
was contacted. Addresses shown are those written into the workbook's own code.*

# Recording the SAP script the macro is built from

## The shortcut

**`Alt` + `F12`** — the "Customize Local Layout" menu on the SAP GUI toolbar.
Then pick **Script Recording and Playback**.

In SAP GUI 7.60 and later the entry sits one level down: `Alt`+`F12` → **More…** →
**Script Recording and Playback**. Either way the dialog is the same, and there is no
direct key combination for it.

In the dialog:

1. **Record** (the red dot) → choose where to save the `.vbs`
2. Do the clicks you want captured
3. **Stop** (the square)

Ctrl-clicking the toolbar's "Customize Local Layout" button gets you to the same menu if
`Alt`+`F12` is intercepted by your terminal or remote-desktop client.

## If Alt+F12 shows no scripting entry

Scripting is switched off. It needs both halves:

- **Server:** profile parameter `sapgui/user_scripting = TRUE` (Basis sets it; `RZ11` shows
  its current value). Read-only check: `RZ11` → parameter name → Display.
- **Client:** SAP GUI Options → *Accessibility & Scripting* → *Scripting* → **Enable
  scripting**. Untick *Notify when a script attaches to SAP GUI* and *Notify when a script
  opens a connection* as well, or every step of the macro waits for a click.

Worth knowing before you ask: enabling scripting is logged, and on most landscapes the
security audit log (`SM19`/`SM20`) records the attach. For audit-support work that is a
feature — it evidences who pulled the data — but tell Basis what you are doing rather than
having them find out from the log.

## What to record

Record **one** payment line, end to end, slowly, in a single session. One good recording of
one line gives every element ID the macro needs for all 56.

Take sample #1 as the walkthrough — Sep-25, 03/09/2025, 8,072,447.42, `ACH PYMTS - LCL BULK
FNDG`:

| # | Do this | What it captures |
|---|---|---|
| 1 | Start the recording from a clean `SESSION_MANAGER` screen | a stable starting point |
| 2 | Enter `FEBAN` | how your release opens the transaction |
| 3 | Type company code `GBKM` | `FEBAN.CompanyCode` |
| 4 | Type statement date `01.09.2025` to `30.09.2025` | `FEBAN.StatementDateFrom` / `…To` |
| 5 | Press Execute | `FEBAN.ExecuteButton` |
| 6 | Click once on the row for 8,072,447.42 | `FEBAN.ResultGrid`, and the grid's real ID |
| 7 | Export the list: right-click → *Spreadsheet…* or the toolbar export button | `Export.*` |
| 8 | In the save dialog, set a folder and file name, press Generate | `Save.*` |
| 9 | Go back, then drill into the posted document (double-click the row, or the *Document* button) | how the drill-down works on your release |
| 10 | In `FB03`, click the *Services for Object* toolbox → *Attachment list* | `FB03.GosToolbox`, `Attach.ListGrid` |
| 11 | Close the attachment list with **Cancel**, not Enter | avoids opening the viewer |
| 12 | Stop the recording | |

Then **run `modFeban.DumpGridColumns`** with a FEBAN result list on screen. It prints the
grid's real column names, which is what `FEBAN.Col.*` needs. Guessing at those is the single
commonest reason a run matches nothing — the example values in the workbook (`VALUT`,
`KWBTR`, `ESTAT`, `BELNR`) are typical but not universal.

## Turning the recording into the Screen Map

Open the `.vbs` in Notepad. It is a list of lines like:

```vbscript
session.findById("wnd[0]/usr/ctxtFEBA_SEL-BUKRS").text = "GBKM"
session.findById("wnd[0]/tbar[1]/btn[8]").press
```

Copy the string inside each `findById("…")` into column F of the **Screen Map** sheet, next
to the matching key. Column D holds an illustrative example for each one — useful for
recognising which line is which, but do not assume it matches your system.

You only need the rows marked *Required = Yes* to get a first run going. The optional ones
add the drill-down and the attachment steps.

## Things the recording will not tell you

- **Your date format.** The recording shows `"01.09.2025"` or `"01/09/2025"` as *you* typed
  it. Put the matching code (`DMY`, `DMY/`, `MDY`, `YMD`) on the Control sheet — `SU3` →
  *Defaults* → *Date format* is the authority. Get it wrong and FEBAN searches a period that
  does not exist and reports nothing found, with no error.
- **Which house bank and account.** The samples appear to span at least two Citibank
  accounts (see `Data Issues`), so leaving both filters blank is the safer default.
- **Batch-input style differences.** A recording made with a different SAP GUI theme
  (Blue Crystal vs Quartz vs Fiori visual theme) can produce different container IDs for the
  same grid. Record on the theme you will run on.

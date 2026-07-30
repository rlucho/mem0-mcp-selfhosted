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

## What the four recordings settled

All four are in [`recordings/`](recordings/), decoded from UTF-16 to plain text. Between them
they cover **steps 1 to 7** of the process, and their IDs are captured values rather than
guesses — marked `recorded` on the Screen Map.

```
FEBAN
  selection popup (wnd[1]!)   GBKM, 01092025 .. 30092025, Execute
  result grid                 wnd[0]/shellcont/shell
  export &MB_EXPORT > &XXL    ctxtDY_PATH + ctxtDY_FILENAME, btn[0]
  double-click the row        AZDAT column
statement item detail         txtD2201_BELNR          -> the FI document
  F2
FI document overview          double-click the AUGBL line
line item detail              txtBSEG-AUGBL           -> the clearing document
  F2
clearing document             mbar/menu[5]/menu[3]    -> Environment > Payment Usage
  mbar/menu[0]/menu[3]/menu[1]                        -> save-list dialog
save dialog                   ctxtDY_PATH, ctxtDY_FILENAME, btn[0]
```

Things worth catching, several of which overturned an earlier guess:

- **The FEBAN selection screen is a modal popup (`wnd[1]`), not a full screen.** Anything
  written against `wnd[0]/usr/...` fails outright. It is also why replaying one of these
  files from the SAP menu errors on line 17 — see the brief at the end.
- **Dates are typed as 8 bare digits** — `01092025`, no separators. `SAP date format` is set
  to `DDMMYYYY` to match. Wrong here is a *silent* failure: FEBAN just returns nothing.
- **The statement-date column is `AZDAT`**, not the `VALUT` that was assumed. Every sample
  match depends on it.
- **The ALV export is `&MB_EXPORT` then `&XXL`**, not `&PC`.
- **The BSEG column to open is `AUGBL`**, the clearing document. `Audit.vbs` used `DMBTR` and
  `Audit2.vbs` `PRCTR` — those were just where the cursor happened to sit.
- **The save dialog confirms with `btn[0]`, not `btn[11]`**, and `ctxtDY_FILENAME` is real.
- **Every export is named `.XLSX`**, so these come back as workbooks, not delimited text.

## What is still missing

**All four recordings stop at step 7**, the Payment Usage export. Nothing reaches FBL1N, the
attachment list, or a PDF.

| Gap | Screen Map keys | State |
|---|---|---|
| Steps 8–9, FBL1N and the batch's ZP payments | `Fbl1n.*`, `MultiSel.*` | PREDICTED — written from standard SAP |
| Step 10, the SCF extra hop | `Scf.OpenInvoices`, `Scf.InvoiceListMenu` | BLOCKED |
| Step 10, the invoice PDF | `Invoice.GosToolbox`, `Invoice.AttachListGrid`, `Invoice.SaveButton` | PREDICTED |

The `Fbl1n.*` rows do not need a recording — `modProbe.ProbeCurrentScreen` checks them against
the live screen, reading only. The step-10 rows are what the brief below is for.


---

# Brief: recording the SCF hop and the PDF download

This is the last gap. Five Screen Map rows come out of it.

## Why the last two recordings looked truncated

`Audit3.vbs` and `Audit5.vbs` are not corrupted, and neither was `Audit2.vbs`. They start at
`wnd[1]/usr/ctxtSL_BUKRS-LOW` — FEBAN's selection popup — because the recorder was started
with that popup already on screen. Replaying one from the SAP menu fails on that line with
"The control could not be found by id", since `wnd[1]` does not exist yet. Nothing is wrong
with the files.

The practical lesson for this recording: **the start point is part of the recording.** Note
where you were when you pressed Record, because that is where playback has to begin.

## Before the real walk, prove the recorder is capturing

The earlier `Audit2.vbs` came back holding only the boilerplate and the `resizeWorkingPane`
line SAP writes the instant recording starts. Two minutes to rule that out:

1. Press **Record**, click any menu, press **Stop**.
2. Open the `.vbs`. If there is no `findById(...)` line *after* `resizeWorkingPane`, the
   recorder is not capturing and nothing you do next will be saved.

## Where to start

Walk a sample by hand as far as **the largest ZP payment of a batch, where that payment is
Santander SCF**. That is steps 1–9. Stop there and start recording.

Note the branch is decided by the largest ZP payment *inside the batch*, not by the party
column in the audit request — so you find it by walking, not by picking a row in advance. Any
of the eight SCF samples is a reasonable place to start looking: #8 Oct, #16 Nov, #24 Jan,
#30 Feb, #34 Mar, #39 Apr, #46 May, #53 Jun.

## What to capture

| # | Do this | Fills |
|---|---|---|
| 1 | From the SCF payment, whatever you click to reach the invoices it settled | `Scf.OpenInvoices` |
| 2 | The menu that lists those invoices | `Scf.InvoiceListMenu` |
| 3 | Export that list — *List ▸ Save/Send ▸ File*, and **type a filename** | reuses `Export.ListMenu` / `Save.*` |
| 4 | Open the largest invoice on it | — |
| 5 | *Services for Object* ▸ *Attachment list* | `Invoice.GosToolbox` |
| 6 | Select the PDF and **Save** it to disk | `Invoice.AttachListGrid`, `Invoice.SaveButton` |
| 7 | Close the attachment list with **Cancel**, not Enter | avoids the viewer |

## The one thing to watch for at step 6

If selecting the attachment opens **Adobe or a Windows viewer window** rather than a SAP
save dialog, stop and say so. That means the PDF is being handed to an external application,
which a script cannot capture — no amount of recording fixes it. The answer then is not more
scripting but either transaction **OAOR** (Business Document Navigator, which does offer a
real export) or a read-only content-server extract from Basis covering all 56 at once.

Worth knowing before you spend time on it: that outcome is common, and it is the single most
likely reason this last step cannot be automated.

## Faster alternative, once the workbook is set up

Recording depends on catching every click. `modProbe.ProbeCurrentScreen` does not — put SAP
on a screen and it reports every mapped ID that is actually there, plus the control type,
reading only. `modProbe.DumpScreen` lists everything on the screen when a prediction misses.
If you have the modules imported already, run the probe at each stop *as well as* recording.

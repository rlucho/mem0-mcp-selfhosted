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

## What the two recordings gave us

Both are in [`recordings/`](recordings/), decoded from UTF-16 to plain text.

**`Audit.vbs` — 38 captured steps, and it covers the main chain.** Everything it touched is
already filled into the Screen Map, marked `recorded` in the Source column. The flow it
proves out:

```
FEBAN
  selection popup (wnd[1]!)   GBKM, 01092025 .. 30092025, Execute
  result grid                 wnd[0]/shellcont/shell, sorted on KWBTR
  double-click the row
statement item detail         txtD2201_BELNR          -> the FI document
  F2
FI document overview          double-click the DMBTR line
line item detail              txtBSEG-AUGBL           -> the clearing document
  F2
clearing document             mbar/menu[5]/menu[3]    -> cleared items + supplier names
  mbar/menu[0]/menu[3]/menu[1]                        -> save-list dialog
save dialog                   ctxtDY_PATH, btn[0]
```

Three things in it were worth catching:

- **The FEBAN selection screen is a modal popup (`wnd[1]`), not a full screen.** Anything
  written against `wnd[0]/usr/...` would fail outright.
- **Dates were typed as 8 bare digits** — `01092025`, no separators. The Control sheet's
  `SAP date format` is set to `DDMMYYYY` to match. This one is a silent failure if wrong:
  FEBAN just returns nothing.
- **The save dialog confirms with `btn[0]`, not `btn[11]`.** And the recording never set
  `DY_FILENAME`, so it kept SAP's default name — that field's ID is the one guess left in the
  save path, and it's marked `VERIFY`.

**`Audit2.vbs` captured nothing.** It holds the scripting boilerplate and a single
`resizeWorkingPane` call — no steps at all. The recorder was almost certainly stopped before
the actions, or started after them. So the Santander SCF confirming-payment route is
**BLOCKED**: those samples run as far as identifying the supplier, then log `BLOCKED_SCF` with
the clearing document and invoice numbers so they can be finished by hand.

## What still needs recording

Three gaps. Everything marked `VERIFY` or `BLOCKED` on the Screen Map:

| Gap | What to record | Screen Map keys |
|---|---|---|
| **The Santander SCF route** | Re-record `Audit2.vbs` from the clearing document through to the saved PDF for a confirming payment | `Scf.Step1..3` |
| **The regular-supplier PDF** | The same walk for a non-Santander supplier: open the payment, *Services for Object* → *Attachment list*, save the PDF | `Invoice.GosToolbox`, `Invoice.AttachListGrid`, `Invoice.SaveButton` |
| **The save dialog's file-name field** | Type a file name into the save dialog this time, rather than accepting the default | `Save.FileName` |

Plus one that isn't a recording: **run `modFeban.DumpGridColumns`** with a FEBAN result list on
screen. `Audit.vbs` only ever touched the `KWBTR` column, so the *value date* column name is
still a guess (`VALUT`), and matching every sample depends on it. This prints the real names.

When you re-record, keep it slow and single-purpose, and **close any attachment list with
Cancel rather than Enter** — Enter hands the document to the external viewer and blocks the
script behind it.

## Turning a recording into the Screen Map

Open the `.vbs` in Notepad — SAP writes it as UTF-16, so it looks fine in Notepad but spaced
out in some editors. It is a list of lines like:

```vbscript
session.findById("wnd[1]/usr/ctxtSL_BUKRS-LOW").text = "gbkm"
session.findById("wnd[1]/tbar[0]/btn[8]").press
```

Copy the string inside each `findById("…")` into **column F** of the **Screen Map** sheet, next
to the matching key. Column F is the only column the macro reads. Column D says where the
current value came from:

| Source | Meaning |
|---|---|
| `recorded` | Came out of `Audit.vbs`. Should be right for this system. |
| `VERIFY` | A standard SAP ID, but the recording never touched it. Check before an EXTRACT run. |
| `BLOCKED` | Unknown, because `Audit2.vbs` captured nothing. |

Ignore the recording's noise lines: `resizeWorkingPane`, `setFocus`, `caretPosition`,
`sendVKey 4` followed by `wnd[2].close` (that's an F4 browse someone opened and cancelled).
None of them are steps.

## Things a recording will not tell you

- **Whether the date format is what SAP expects.** `Audit.vbs` typed `01092025`, so
  `DDMMYYYY` is known to work here. If you switch it, check `SU3` → *Defaults* →
  *Date format* first. A wrong value means FEBAN searches a period that does not exist and
  reports nothing found, with no error at all.
- **Which house bank and account.** The samples appear to span at least two Citibank
  accounts (see `Data Issues`), so leaving both filters blank is the safer default.
- **What an unlabelled button does.** `Audit.vbs` presses `wnd[0]/tbar[1]/btn[14]` right after
  Execute. That is in the Screen Map as `FEBAN.PostExecuteButton` and it is optional — check
  what it actually does before relying on it, or clear the cell to skip it.
- **Theme differences.** A recording made under a different SAP GUI theme (Blue Crystal vs
  Quartz vs the Fiori visual theme) can produce different container IDs for the same grid.
  Record on the theme you will run on.

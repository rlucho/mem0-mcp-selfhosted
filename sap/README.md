# SAP PER → FBL1N attach test

A minimal connectivity test: attach to the SAP GUI session for system **PER**
that is *already open* on the PC, and navigate it to **FBL1N** (Vendor Line Item
Display). Nothing is executed, posted or exported — it stops on the FBL1N
selection screen.

Two equivalent implementations, same SAP COM API, different host process:

| File | Host | Run it with |
|---|---|---|
| `modSapFbl1nTest.bas` | `EXCEL.EXE` | VBE → File → Import File…, then Alt+F8 → `Test_SAP_PER_FBL1N` |
| `SAP_PER_FBL1N_Test.vbs` | `wscript.exe` / `cscript.exe` | double-click, or `cscript.exe SAP_PER_FBL1N_Test.vbs /nologo` |

Import the `.bas` — do not paste its contents into a module. The first line
(`Attribute VB_Name = …`) is only valid at import time and is a compile error if
typed into the code pane.

Both print the same report, and both save a copy to
`SAP_PER_FBL1N_Test.txt` (next to the script for the `.vbs`, in `%TEMP%` for the
Excel module). The `.vbs` exits `0` on pass, `1` on fail.

## Prerequisites

- SAP GUI already **running and logged on to PER**. Neither file logs on;
  `GetObject()` only attaches to a window that is already there.
- Scripting enabled **client side** (Options → Accessibility & Scripting →
  Scripting → Enable scripting) **and server side** (`sapgui/user_scripting`).
- Excel/wscript and SAP GUI at the **same elevation** — COM will not attach
  across a normal ↔ "run as administrator" boundary.

## Sample output

```
SAP PER -> FBL1N connectivity test
----------------------------------------------------------
[1] attaching to the running SAP GUI
    moniker SAPGUISERVER -> not registered (error -2147221020)
    moniker SAPGUI -> OK
    connections open: 2
    conn(0) "E01 Production"  sessions=1
      sess(0) System=E01  Client=100  User=BFERRO1  Tx=SESSION_MANAGER
    conn(1) "PER Production"  sessions=1
      sess(0) System=PER  Client=100  User=BFERRO1  Tx=SESSION_MANAGER
    selected: the PER session (idle, no popup)
    attached - transaction before = SESSION_MANAGER
[2] sending /nFBL1N to the command field
[3] result
    transaction = FBL1N
    window      = Vendor Line Item Display

PASS - attached to PER and FBL1N is on screen.
```

## Reading a failure

Every step logs its own diagnosis; the inventory of open sessions is always
printed, so a failure is self-explanatory.

- **Both monikers "not registered"** — SAP GUI is not open, scripting is off, or
  the elevation differs. `-2147221020` is `MK_E_SYNTAX`: the *name* did not
  resolve. It is not a security block, even though VBA reports it as
  "Automation error / Invalid syntax".
- **"no SAP connection is open"** — the engine answered but nothing is logged on.
- **"No open session reports SystemName = PER"** — PER is not open. The listed
  sessions show what *is*; if PER's SID is spelled differently, change
  `TARGET_SYSTEM` at the top of the file. Nothing is sent in this case: driving
  a system that was not asked for is the failure mode this test exists to avoid.
- **"busy or has a popup open"** — a modal dialog is up, or a long-running step
  is in flight. The session is deliberately left untouched.
- **Landed somewhere other than FBL1N** — read the status bar line in the report
  (missing authorisation, or the previous screen refused to leave).

## Design notes

Four things that are easy to get wrong and cost real debugging time:

1. **Two COM monikers exist.** Classic SAP GUI (`saplogon.exe`) publishes
   `SAPGUI`; SAP Business Client / NWBC publishes `SAPGUISERVER`. A machine
   answers to one and fails the other with `MK_E_SYNTAX`. Hardcoding either name
   breaks half a mixed fleet, so both are tried and the winner is cached.
2. **Never index a SAP GUI collection with a `Long`.** `app.Children(0)` works
   only because the literal is typed `Integer` (VT_I2); a `Long` loop counter
   raises error 618, "Bad index type for collection access". Hence `CInt()` on
   every collection index. This is universal to SAP GUI, not an NWBC quirk.
3. **The session is picked by `SystemName`**, not by the customary
   `Children(0).Children(0)`. Index 0 is whichever connection was opened first,
   so with two systems open the classic pattern silently drives the wrong one.
   A readable `SystemName` is authoritative; the connection *description* is used
   only as a fallback when `SystemName` cannot be read at all.
4. **`On Error Resume Next` has a sharp edge.** A property read that throws while
   an argument is being evaluated abandons the *entire* statement — the log line
   vanishes rather than showing `<unreadable>`. And `For j = 0 To coll.Count - 1`
   with a throwing limit expression enters the body **once** with `j`
   uninitialised. So risky properties are read one per statement into their own
   variable, and collection counts are captured *before* the `For`.

5. **`Enum` is a reserved word — in both languages.** A local named `eNum` (an
   obvious shorthand for "error number") *is* the keyword `Enum`, because both
   languages are case-insensitive. It costs a bare `Compile error: Syntax error`
   on the `Dim` line. Imported code keeps its original casing, so the VBE never
   recapitalises it to `Enum` and gives the game away — the variable just looks
   fine. VBScript reserves `Enum` too, even though it has no `Enum` statement.
   Hence `errNum`. Same trap waits behind `Date`, `Error`, `Len`, `Line`,
   `Name`, `String`, `Text`, `Type` and `Module`.

Also: keep the `.vbs` saved **without a BOM** — `cscript` rejects a UTF-8 BOM
with "Invalid character" at (1,1). `.gitattributes` checks both files out with
CRLF endings.

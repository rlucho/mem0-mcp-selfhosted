import os, re, glob

def logical(path):
    out, buf, start = [], "", None
    for n, raw in enumerate(open(path, encoding="utf-8"), 1):
        s = raw.rstrip("\n")
        if start is None: start = n
        c, q = "", False
        for ch in s:
            if ch == '"': q = not q
            if ch == "'" and not q: break
            c += ch
        c = c.rstrip()
        if c.endswith("_"):
            buf += c[:-1]; continue
        out.append((start, (buf + c).strip())); buf, start = "", None
    return out

# Resolve the module folder rather than trusting the cwd. glob("*.bas") in
# the repo root matches nothing and every rule then passes vacuously -- which
# is exactly how a Public Const placed after a procedure reached a live run.
_here = os.path.dirname(os.path.abspath(__file__))
_candidates = [os.getcwd(),
               os.path.join(os.getcwd(), "sap-audit-macro", "vba"),
               os.path.join(os.getcwd(), "vba"),
               "/home/user/mem0-mcp-selfhosted/sap-audit-macro/vba"]
_dir = next((d for d in _candidates if glob.glob(os.path.join(d, "*.bas"))), None)
if _dir is None:
    raise SystemExit("no .bas files found -- refusing to report a vacuous pass")
os.chdir(_dir)
files = sorted(glob.glob("*.bas"))
print(f"checking {len(files)} modules in {_dir}")
modules = {p[:-4].lower() for p in files}
udts, procs, members, modvars, consts = set(), set(), set(), set(), set()
for path in files:
    src = open(path, encoding="utf-8").read()
    for m in re.finditer(r'^\s*(?:Public |Private )?Type\s+(\w+)', src, re.M): udts.add(m.group(1).lower())
    for m in re.finditer(r'^\s*(?:Public |Private )?(?:Sub|Function|Property\s+\w+)\s+(\w+)', src, re.M): procs.add(m.group(1).lower())
    for blk in re.findall(r'Type\s+\w+(.*?)End Type', src, re.S):
        for m in re.finditer(r'^\s*(\w+)\s+As\s', blk, re.M): members.add(m.group(1).lower())
    for m in re.finditer(r'^(?:Public|Private)\s+(?:Const\s+)?(\w+)\s*(?:\([^)]*\))?\s+As\s', src, re.M): modvars.add(m.group(1).lower())
    for m in re.finditer(r'^(?:Public|Private)\s+Const\s+(\w+)', src, re.M): consts.add(m.group(1).lower())
    for m in re.finditer(r'Declare\s+(?:PtrSafe\s+)?(?:Sub|Function)\s+(\w+)', src): procs.add(m.group(1).lower())

VBA = set(w.lower() for w in """
If Then Else ElseIf End Sub Function Dim Set Let Const Static For Each Next To Step While
Wend Do Loop Until Select Case Exit On Error GoTo Resume Public Private As New Nothing
True False Empty Null And Or Not Xor Mod Is In Like ByVal ByRef Optional ParamArray Call
With Type Enum Property Get ReDim Preserve Erase Option Explicit Attribute Declare Lib
Alias PtrSafe String Long Integer Double Boolean Date Variant Object Currency Byte Single
Collection Workbook Worksheet Worksheets Range Application ThisWorkbook Me Debug Print Err
Now Time Format Trim Len Left Right Mid InStr InStrRev Replace Split Join UCase LCase CStr
CLng CDbl CDate CBool CInt Abs Int Fix Val IsNumeric IsDate IsObject IsEmpty IsNull IsArray
Array LBound UBound Open Close Input Output Binary Access Read Write Put FreeFile LOF Seek
Kill MsgBox InputBox DoEvents RGB Chr Asc Space IIf Sgn Sqr TypeName VarType CreateObject
GetObject Environ Dir Shell Randomize Rnd Timer DateSerial DateAdd DateDiff Year Month Day
Hour Minute Second Weekday StrComp StrConv StrReverse LTrim RTrim Hex Oct Cells Rows Columns
Shapes Buttons OLEObjects Names Sheets Count Item Add Remove Delete Select Activate Value
Text Name Path Size Length Number Description Source Raise Clear Keys Exists CompareMode
Visible Font Color Bold Interior Caption OnAction Top Width Height Placement Characters
vbNullString vbCrLf vbCr vbLf vbTab vbYes vbNo vbOK vbCancel vbYesNo vbOKCancel vbQuestion
vbInformation vbExclamation vbCritical vbDefaultButton1 vbDefaultButton2 vbTextCompare
vbBinaryCompare vbObjectError VbMsgBoxResult xlUp xlWhole xlSheetVeryHidden xlValidateList
xlValidAlertStop xlBetween xlCenter Formula1 Operator AlertStyle
xlOpenXMLWorkbookMacroEnabled VBA7 Win64 SaveChanges fileName UpdateLinks ReadOnly What
xlTop xlOpenXMLWorkbook FileFormat After Before Link DisplayAsIcon IconLabel
IconFileName Anchor Address TextToDisplay SubAddress ScreenTip Left Top Width Height
Destination Origin Filename fileFilter Title vbBoolean vbDate vbString vbEmpty vbNull vbYesNoCancel vbCancel UpdateLinks ButtonName
LookAt MatchCase Offset Find End Rows Columns Copy Paste
""".split())

known = VBA | modules | udts | procs | members | modvars | consts
problems = []

for path in files:
    proc, local, labels, body = None, set(), set(), []
    for n, s in logical(path):
        m = re.match(r'(?:Public |Private )?(?:Sub|Function)\s+(\w+)\s*\((.*)\)', s)
        if m:
            proc, local, labels, body = m.group(1), set(), set(), []
            local.add(m.group(1).lower())
            for pm in re.finditer(r'(\w+)(?:\(\))?\s+As\s+\w+', m.group(2)): local.add(pm.group(1).lower())
            continue
        if re.match(r'End (Sub|Function)', s):
            if proc:
                for ln, tok in body:
                    t = tok.lower()
                    if t not in known and t not in local and t not in labels:
                        problems.append(f"{path}:{ln}  {proc}(): '{tok}'")
            proc = None
            continue
        if proc is None: continue
        if re.match(r'^\w+:$', s): labels.add(s[:-1].lower()); continue
        # declarations, including multi-variable Dim lines
        d = re.match(r'(?:Dim|Static|Const|ReDim(?:\s+Preserve)?)\s+(.*)$', s)
        if d:
            for nm in re.finditer(r'(\w+)\s*(?:\([^)]*\))?\s+As\s+\w+', d.group(1)): local.add(nm.group(1).lower())
            for nm in re.finditer(r'(?:^|,)\s*(\w+)\s*(?:\([^)]*\))?\s*(?:,|$)', d.group(1)): local.add(nm.group(1).lower())
        for g in re.finditer(r'(?:GoTo|Resume)\s+(\w+)', s): labels.add(g.group(1).lower())
        stripped = re.sub(r'"[^"]*"', '""', s)
        stripped = re.sub(r'#\w+', '', stripped)
        for tm in re.finditer(r'(?<![\w.$#])([A-Za-z_]\w*)', stripped):
            body.append((n, tm.group(1)))

seen, out = set(), []
for p in problems:
    k = p.split("  ")[1]
    if k in seen: continue
    seen.add(k); out.append(p)
print("\n".join(out) if out else "clean: every identifier resolves to a declaration, label, module, type or VBA builtin")

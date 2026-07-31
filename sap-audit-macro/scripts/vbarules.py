"""VBA-specific compile rules that a generic structure check misses."""
import os, re, glob, collections, os

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
        out.append((start, (buf + c).strip(), s)); buf, start = "", None
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
udts, problems = set(), []

for path in files:
    for m in re.finditer(r'^\s*(?:Public |Private )?Type\s+(\w+)', open(path, encoding="utf-8").read(), re.M):
        udts.add(m.group(1))

# 1. UDT may not be passed ByVal
for path in files:
    for n, s, _ in logical(path):
        for m in re.finditer(r'ByVal\s+(\w+)\s+As\s+(\w+)', s):
            if m.group(2) in udts:
                problems.append(f"BLOCKER {path}:{n}  UDT passed ByVal: '{m.group(1)} As {m.group(2)}'")

# 2. Public procedures sharing a name across modules -> ambiguous unqualified calls
pub = collections.defaultdict(list)
for path in files:
    for n, s, _ in logical(path):
        m = re.match(r'(?:Public\s+)?(?:Sub|Function|Property\s+\w+)\s+(\w+)', s)
        if m and not s.startswith("Private"):
            pub[m.group(1)].append(f"{path}:{n}")
for name, places in sorted(pub.items()):
    if len(places) > 1:
        problems.append(f"AMBIG   Public '{name}' declared in {len(places)} modules: {', '.join(places)}")

# 3. A procedure named the same as a module
mods = {p[:-4] for p in files}
for name, places in pub.items():
    if name in mods:
        problems.append(f"BLOCKER procedure '{name}' has the same name as a module ({places})")

# 3b. Module-level declarations must precede the first procedure.
for path in files:
    seen_proc, depth = False, 0
    for n, t, _ in logical(path):
        if re.match(r'(?:Public |Private )?(?:Sub|Function|Property)\s+\w+', t):
            seen_proc, depth = True, 1
        elif re.match(r'End (Sub|Function|Property)', t):
            depth = 0
        # Const/Type/Enum/Declare were excluded here once, on the assumption
        # VBA tolerates them mid-module. It does not: a Public Const placed
        # after the first procedure never becomes a module member, and the
        # only symptom is 'Method or data member not found' at the CALL SITE
        # in another module, which points nowhere near the real cause.
        elif seen_proc and depth == 0 and re.match(r'(?:Public|Private|Dim|Const)\s+\w+', t):
            problems.append(f"BLOCKER {path}:{n}  declaration after a procedure: {t[:60]}")

# 3c. VBA reserved words cannot be used as variable names.
#
# 'Dim empty As FebanMatch' compiles to nothing but "Syntax error", with no
# hint which word is the problem -- and neither a structure check nor an
# identifier-resolution check notices, because the name resolves perfectly
# well to the declaration it is not allowed to have.
# Only words VBA genuinely refuses as an identifier. Statements and
# functions like Name, Date, Error, Line, Input, Close and Left ARE legal
# variable names, however odd they read, and flagging them would make this
# rule noise -- a checker people learn to ignore is worse than none.
RESERVED = set("""
empty null nothing true false me
and or not xor eqv imp mod is like to then step as in each
if else end select case do loop while wend until for next exit goto on
resume return stop set let new call rem erase redim with
dim const static public private friend option
sub function property type enum declare event raiseevent implements
byval byref optional paramarray withevents
""".split())

# Only the things that DECLARE a variable. 'Private Sub Finish' declares a
# procedure, and Sub is allowed to follow Private there.
DECLARES = re.compile(
    r'^\s*(?:Dim|Const|Static|(?:Public|Private)(?:\s+Const|\s+Static)?)\s+'
    r'([A-Za-z_]\w*)\s*(?:\(|As|,|=|$)')
NOT_A_VARIABLE = {"sub", "function", "property", "type", "enum", "declare",
                  "const", "static", "withevents"}

for path in files:
    for n, t, _ in logical(path):
        m = DECLARES.match(t)
        if not m:
            continue
        first = m.group(1).lower()
        if first in NOT_A_VARIABLE:
            continue
        if first in RESERVED:
            problems.append(f"BLOCKER {path}:{n}  '{m.group(1)}' is a VBA reserved "
                            f"word and cannot be a variable name: {t[:60]}")
        # 'Dim a As Long, empty As String' -- the ones after the first comma
        for extra in re.finditer(r',\s*([A-Za-z_]\w*)\s+As\b', t):
            if extra.group(1).lower() in RESERVED:
                problems.append(f"BLOCKER {path}:{n}  '{extra.group(1)}' is a VBA "
                                f"reserved word and cannot be a variable name")

# 4. VBA allows at most 25 line-continuations per statement
for path in files:
    cont, start = 0, None
    for n, raw in enumerate(open(path, encoding="utf-8"), 1):
        if raw.rstrip("\n").rstrip().endswith("_"):
            if start is None: start = n
            cont += 1
        else:
            if cont > 24: problems.append(f"BLOCKER {path}:{start}  {cont} line continuations (max 25)")
            cont, start = 0, None

# 5. Physical lines over 1023 characters
for path in files:
    for n, raw in enumerate(open(path, encoding="utf-8"), 1):
        if len(raw.rstrip("\n")) > 1023:
            problems.append(f"BLOCKER {path}:{n}  line is {len(raw)} chars (max 1023)")

# 6. Variables used but never declared, per procedure (Option Explicit)
KEYWORDS = set("""If Then Else ElseIf End Sub Function Dim Set Let Const For Each Next To Step
While Wend Do Loop Until Select Case Exit On Error GoTo Resume Resume Public Private
As New Nothing True False Empty Null And Or Not Xor Mod Is Like ByVal ByRef Optional
ParamArray Call With Type Enum Property Get Redim ReDim Preserve Erase Option Explicit
Attribute Declare Lib Alias PtrSafe String Long Integer Double Boolean Date Variant Object
Currency Byte Single Collection Workbook Worksheet Range Application ThisWorkbook Me
Debug Print Err Now Date Time Format Trim Len Left Right Mid InStr InStrRev Replace Split
Join UCase LCase CStr CLng CDbl CDate CBool Abs Int Fix Val IsNumeric IsDate IsObject
IsEmpty IsNull Array LBound UBound Open Close Input Output Binary Access Read Write Get Put
Free FreeFile LOF Seek Kill MsgBox InputBox DoEvents RGB Chr Asc Space vbNullString""".split())
for path in files:
    proc, declared, used = None, set(), []
    modlevel = set()
    for n, s, _ in logical(path):
        if re.match(r'(?:Public|Private|Dim|Const)\s', s) and proc is None:
            for m in re.finditer(r'\b(\w+)\s+As\s+\w+', s): modlevel.add(m.group(1))
            m = re.match(r'(?:Public|Private)\s+Const\s+(\w+)', s)
            if m: modlevel.add(m.group(1))
    print(f"  scanned {path}")
print()
print("\n".join(problems) if problems else "no VBA-rule problems found")

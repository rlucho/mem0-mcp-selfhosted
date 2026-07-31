import os, re, glob, collections, sys

def logical_lines(path):
    out, buf, start = [], "", None
    for n, raw in enumerate(open(path, encoding="utf-8"), 1):
        s = raw.rstrip("\n")
        if start is None: start = n
        clean, in_str = "", False
        for c in s:
            if c == '"': in_str = not in_str
            if c == "'" and not in_str: break
            clean += c
        clean = clean.rstrip()
        if clean.endswith("_"):
            buf += clean[:-1]; continue
        out.append((start, (buf + clean).strip())); buf, start = "", None
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
public = collections.defaultdict(set)
types  = {}
for path in files:
    mod = path[:-4]
    for n, s in logical_lines(path):
        m = re.match(r'Public\s+(?:Sub|Function)\s+(\w+)', s) or re.match(r'^(?:Sub|Function)\s+(\w+)', s)
        if m: public[mod].add(m.group(1))
        m = re.match(r'Public\s+Const\s+(\w+)', s)
        if m: public[mod].add(m.group(1))
        m = re.match(r'Public\s+Type\s+(\w+)', s)
        if m: public[mod].add(m.group(1)); types[m.group(1)] = mod
        m = re.match(r'Public\s+(\w+)\s+As\s', s)
        if m: public[mod].add(m.group(1))

problems = []
# cross-module refs
for path in files:
    for n, s in logical_lines(path):
        for mod, member in re.findall(r'\b(mod[A-Z]\w+)\.(\w+)', s):
            if re.search(r'"[^"]*' + mod + r'\.' + member, s): continue   # inside a string
            if mod not in public: problems.append(f"{path}:{n} unknown module {mod}")
            elif member not in public[mod]:
                problems.append(f"{path}:{n} {mod}.{member} is not Public")

# UDTs used as declarations must be Public Types
for path in files:
    mod = path[:-4]
    for n, s in logical_lines(path):
        m = re.match(r'(?:Dim|Public|Private)\s+\w+\s+As\s+(\w+)$', s) or \
            re.search(r'\)\s+As\s+(\w+)$', s)
        if m:
            t = m.group(1)
            if t in types and types[t] != mod and t not in public[types[t]]:
                problems.append(f"{path}:{n} type {t} not public")

# block balance
for path in files:
    depth = collections.Counter(); proc=None
    for n, s in logical_lines(path):
        if not s: continue
        m = re.match(r'(?:Public |Private )?(?:Sub|Function) (\w+)', s)
        if m: proc=(n,m.group(1)); depth.clear()
        if re.match(r'End (Sub|Function)\b', s):
            for k,v in depth.items():
                if v: problems.append(f"{path} {proc[1]} (line {proc[0]}): {k} {v:+d}")
            proc=None
        if re.match(r'If\b', s) and re.search(r'\bThen$', s): depth["If"]+=1
        if re.match(r'End If\b', s): depth["If"]-=1
        if re.match(r'For\b', s): depth["For"]+=1
        if re.match(r'Next\b', s): depth["For"]-=1
        if re.match(r'Do\b', s): depth["Do"]+=1
        if re.match(r'Loop\b', s): depth["Do"]-=1
        if re.match(r'Select Case\b', s): depth["Select"]+=1
        if re.match(r'End Select\b', s): depth["Select"]-=1
        if re.match(r'With\b', s): depth["With"]+=1
        if re.match(r'End With\b', s): depth["With"]-=1

# GoTo labels
for path in files:
    lines = logical_lines(path); proc=None
    for idx,(n,s) in enumerate(lines):
        if re.match(r'(?:Public |Private )?(?:Sub|Function) \w+', s): proc=idx
        if re.match(r'End (Sub|Function)\b', s) and proc is not None:
            chunk=[t for _,t in lines[proc:idx+1]]
            targets={t[:-1] for t in chunk if re.match(r'^\w+:$',t)}
            for t in chunk:
                if "On Error Resume Next" in t: continue
                for g in re.findall(r'\bGoTo (\w+)',t)+re.findall(r'\bResume (\w+)',t):
                    if g not in targets and g!="0":
                        problems.append(f"{path}:{lines[proc][0]} GoTo {g} unresolved")
            proc=None

print("\n".join(problems) if problems else "OK: cross-module refs, UDTs, blocks and GoTo labels all resolve")
print()
for path in files:
    txt=open(path,encoding="utf-8").read()
    print(f"  {path:20} {len(txt.splitlines()):4} lines  Option Explicit: {'yes' if 'Option Explicit' in txt else 'NO'}")

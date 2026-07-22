import re, os

dump = open("vba_dump.txt", encoding="utf-8", errors="replace").read().splitlines()

# We only care about the real code (before the binary FORM STRING garbage).
# Cut at first "VBA FORM STRING" line.
cut = next((i for i,l in enumerate(dump) if l.startswith("VBA FORM STRING")), len(dump))
dump = dump[:cut]

def is_solid(l):  # section delimiter: a run of dashes, no spaces
    s=l.strip()
    return len(s)>=20 and set(s)=={"-"}
def is_spaced(l): # header/code separator: "- - - - -"
    return "- - -" in l

modules={}
i=0
name=None
while i < len(dump):
    l=dump[i]
    m=re.match(r"VBA MACRO (\S+)", l)
    if m:
        name=m.group(1)
        # advance past "in file:" line and the spaced-dash separator
        i+=1
        while i<len(dump) and not is_spaced(dump[i]):
            i+=1
        i+=1  # skip the spaced-dash line
        code=[]
        while i<len(dump) and not is_solid(dump[i]):
            code.append(dump[i]); i+=1
        txt="\n".join(code).rstrip("\n")
        if txt.strip() and txt.strip()!="(empty macro)":
            modules[name]=txt
        continue
    i+=1

os.makedirs("original_modules", exist_ok=True)
for n,t in modules.items():
    open(f"original_modules/{n}", "w", encoding="utf-8").write(t + "\n")

print("Extracted modules with code:")
for n,t in sorted(modules.items()):
    print(f"  {n:28s} {len(t.splitlines()):5d} lines")

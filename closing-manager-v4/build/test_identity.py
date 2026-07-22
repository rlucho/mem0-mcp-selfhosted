import olefile
from cfb import Cfb, build_cfb

orig = open("extracted/xl/vbaProject.bin","rb").read()
tpl = Cfb(orig)

# rebuild with NO edits
rebuilt = build_cfb(tpl, {})
open("rebuilt_identity.bin","wb").write(rebuilt)

# compare every stream via olefile (authoritative reader)
o1 = olefile.OleFileIO(orig)
o2 = olefile.OleFileIO(rebuilt)
s1 = sorted("/".join(p) for p in o1.listdir(streams=True, storages=False))
s2 = sorted("/".join(p) for p in o2.listdir(streams=True, storages=False))
assert s1 == s2, ("stream set differs", set(s1)^set(s2))
bad = 0
for p in s1:
    d1 = o1.openstream(p).read(); d2 = o2.openstream(p).read()
    if d1 != d2:
        bad += 1; print("  MISMATCH", p, len(d1), len(d2))
o1.close(); o2.close()
print(f"streams compared: {len(s1)}  mismatches: {bad}")
print("file sizes  orig=%d rebuilt=%d" % (len(orig), len(rebuilt)))

# olevba must extract identical code from both
from oletools.olevba import VBA_Parser
def codes(path_or_bytes):
    vp = VBA_Parser("x.bin", data=path_or_bytes)
    out = {}
    for _,_,name,code in vp.extract_macros():
        out[name] = code
    vp.close(); return out
c1 = codes(orig); c2 = codes(rebuilt)
diff = [k for k in c1 if c1[k]!=c2.get(k)]
print("olevba modules:", len(c1), "code diffs:", diff)
print("IDENTITY_OK" if bad==0 and not diff and s1==s2 else "IDENTITY_FAIL")

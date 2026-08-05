"""Simulate the SHIPPED CM_ToAmount against what the ORIGINAL VBA did, for every
value shape SAP can emit. Flags any case where V4 continues but the original
stopped (safety loss), or where V4 returns a different number (silent error)."""
def shipped(v):
    s=str(v).strip(); neg=False
    if s=="": return ("OK",0.0)
    if s.endswith("-"): neg=True; s=s[:-1].strip()
    if s.startswith("-"): neg=True; s=s[1:].strip()
    s=s.replace(" ","").replace("\xa0","")
    pDot=s.rfind("."); pCom=s.rfind(",")
    if pDot>pCom: s=s.replace(",","")
    elif pCom>pDot: s=s.replace(".","").replace(",",".")
    if not s: return ("STOP",None)
    dots=0
    for c in s:
        if c==".":
            dots+=1
            if dots>1: return ("STOP",None)
        elif not c.isdigit(): return ("STOP",None)
    try: x=float(s)
    except: return ("STOP",None)
    return ("OK", -x if neg else x)

def original(v, loc):
    """VBA Round()/CDbl() on a String, under Windows locale `loc`."""
    s=str(v).strip(); neg=False
    if s=="": return ("STOP",None)          # Round("",2) -> error 13
    if s.endswith("-"): neg=True; s=s[:-1].strip()   # macro strips it itself
    dec, grp = ('.',',') if loc=='en' else (',','.')
    t=s.replace(grp,"")
    if t.count(dec)>1: return ("STOP",None)
    t=t.replace(dec,".")
    try: x=float(t)
    except: return ("STOP",None)
    return ("OK", -x if neg else x)

CASES=["1234.56","1234,56","1.234,56","1,234.56","1234.56-","-1234,56",
       "1.234","1,234","1.234.567","1,234,567","1.234.567,89","1,234,567.89",
       "","0","*****","1 234,56","  12,00-  ","0,00","12","1234567"]

print("%-16s | %-14s | %-14s | %-14s | verdict" % ("value","orig en-US","orig de-DE","V4 shipped"))
print("-"*84)
issues=[]
for c in CASES:
    e=original(c,'en'); d=original(c,'de'); v=shipped(c)
    note=""
    # safety loss: BOTH locales stopped, V4 continues
    if e[0]=="STOP" and d[0]=="STOP" and v[0]=="OK":
        note="SAFETY LOSS - used to stop, now returns %s" % v[1]; issues.append((c,note))
    # silent divergence: a locale that worked now gives a different number
    else:
        for lab,o in (("en-US",e),("de-DE",d)):
            if o[0]=="OK" and v[0]=="OK" and abs(o[1]-v[1])>1e-9:
                note="DIVERGES from %s: %s vs %s" % (lab,o[1],v[1]); issues.append((c,note)); break
            if o[0]=="OK" and v[0]=="STOP":
                note="NEW STOP (was fine on %s: %s)" % (lab,o[1]); issues.append((c,note)); break
    print("%-16r | %-14s | %-14s | %-14s | %s" % (c,
        "%s %s"%e if e[0]=="OK" else "STOP",
        "%s %s"%d if d[0]=="OK" else "STOP",
        "%s %s"%v if v[0]=="OK" else "STOP", note))
print()
print("PROBLEMS FOUND:", len(issues))
for c,n in issues: print("  %-16r %s" % (c,n))

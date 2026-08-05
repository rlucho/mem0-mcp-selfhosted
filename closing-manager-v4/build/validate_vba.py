#!/usr/bin/env python3
import io,os,re,sys
SRC=os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),'build_v4')
PROC=re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+)?(?:Static\s+)?(Sub|Function|Property\s+(?:Get|Let|Set))\s+[A-Za-z_]',re.I)
END =re.compile(r'^\s*End\s+(Sub|Function|Property)\b',re.I)
DECL=re.compile(r'^\s*(Public|Private|Global|Dim|Const|Option|Type|Enum|Declare)\b',re.I)
bad=0
for f in sorted(os.listdir(SRC)):
    if not f.endswith('.bas'): continue
    L=io.open(os.path.join(SRC,f),encoding='utf-8',newline='').read().replace('\r\n','\n').split('\n')
    # strip comments + continuations for structure analysis
    procs=ends=0; first=None; late=[]; labels={}; gotos={}
    inproc=False; cont=False
    for i,raw in enumerate(L):
        t=raw.strip()
        prevcont=cont
        cont = t.endswith(' _')
        if prevcont: continue
        if not t or t.startswith("'"): continue
        if PROC.match(t) and not re.match(r'^\s*End\b',t):
            procs+=1; inproc=True
            if first is None: first=i+1
            labels.setdefault(i,None); cur=i
            continue
        if END.match(t):
            ends+=1; inproc=False; continue
        if not inproc and DECL.match(t) and first is not None:
            late.append(i+1)
        if inproc:
            m=re.match(r'^([A-Za-z_]\w*):\s*(?:\'.*)?$',t)
            if m: labels.setdefault('L',set()).add(m.group(1))
            for g in re.findall(r'\bGoTo\s+([A-Za-z_]\w*)',t,re.I): gotos.setdefault('G',set()).add(g)
    miss = (gotos.get('G',set()) - labels.get('L',set())) - {'0'}
    status = 'OK ' if (procs==ends and not late and not miss) else 'BAD'
    if status=='BAD': bad+=1
    print('%s %-26s procs=%-3d ends=%-3d late-decls=%s missing-labels=%s' % (status,f,procs,ends,late or '-',sorted(miss) or '-'))
sys.exit(1 if bad else 0)

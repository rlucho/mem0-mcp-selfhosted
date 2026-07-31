import re, glob, openpyxl, collections

# Keys the VBA asks for
asked = collections.defaultdict(list)
for path in sorted(glob.glob("sap-audit-macro/vba/*.bas")):
    for n, line in enumerate(open(path, encoding="utf-8"), 1):
        if line.strip().startswith("'"): continue
        for m in re.finditer(r'ElementId(?:OrBlank)?\("([^"]+)"\)', line):
            asked[m.group(1)].append(f"{path.split('/')[-1]}:{n}")

# Keys the workbook defines
wb = openpyxl.load_workbook("sap-audit-macro/FEBAN_Audit_Control.xlsx")
sm = wb["Screen Map"]
defined = {}
for r in range(6, 80):
    k = sm.cell(r, 2).value
    if k and "." in str(k):
        defined[str(k).strip()] = (sm.cell(r, 5).value, sm.cell(r, 6).value)

print("=== Screen Map keys ===")
missing = sorted(set(asked) - set(defined))
unused  = sorted(set(defined) - set(asked))
print(f"asked by code: {len(asked)}   defined in workbook: {len(defined)}")
print(f"  asked but NOT defined : {missing if missing else 'none'}")
print(f"  defined but not asked : {unused if unused else 'none'}")

print("\n  required keys and whether they have a value:")
for k, (req, val) in defined.items():
    if req == "Yes":
        print(f"    {'OK ' if val else 'EMPTY'}  {k}")

# Settings the VBA asks for vs the Control sheet
asked_s = collections.defaultdict(list)
for path in sorted(glob.glob("sap-audit-macro/vba/*.bas")):
    for n, line in enumerate(open(path, encoding="utf-8"), 1):
        if line.strip().startswith("'"): continue
        for m in re.finditer(r'Setting(?:Number|IsYes)?\("([^"]+)"', line):
            asked_s[m.group(1)].append(f"{path.split('/')[-1]}:{n}")
# also the caption settings passed as literals into LargestRowWithCaptions
for path in sorted(glob.glob("sap-audit-macro/vba/*.bas")):
    for n, line in enumerate(open(path, encoding="utf-8"), 1):
        for m in re.finditer(r'"((?:SCF invoice list|Cleared list) [a-z ]+column)"', line):
            asked_s[m.group(1)].append(f"{path.split('/')[-1]}:{n}")

ct = wb["Control"]
defined_s = {str(ct.cell(r,2).value).strip() for r in range(13,45) if ct.cell(r,2).value}
print("\n=== Control settings ===")
ms = sorted(set(asked_s) - defined_s)
print(f"asked by code: {len(asked_s)}   defined on sheet: {len(defined_s)}")
print(f"  asked but NOT on the Control sheet: {ms if ms else 'none'}")
for k in ms:
    print(f"      {k}  <- {asked_s[k]}")

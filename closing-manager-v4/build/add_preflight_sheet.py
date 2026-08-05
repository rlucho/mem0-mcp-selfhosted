"""Add a 'Preflight' worksheet with a button bound to the PreflightCheck macro.

Surgical: only adds new parts and makes three tiny edits (workbook.xml,
workbook.xml.rels, [Content_Types].xml). Every existing part is copied through
byte-for-byte, so the START-sheet buttons, forms and vbaProject are untouched.
The macro name is NOT changed - the shape simply calls [0]!PreflightCheck,
the same convention the workbook already uses for [0]!UpdateData etc.
"""
import re, zipfile, shutil, sys, pathlib

SRC = "wb.xlsm"
OUT = "Closing_Manager_IP_V4-CIO.xlsm"

zin = zipfile.ZipFile(SRC)
names = zin.namelist()

# ---- pick free ids -------------------------------------------------------
wb = zin.read("xl/workbook.xml").decode("utf-8")
rels = zin.read("xl/_rels/workbook.xml.rels").decode("utf-8")
ct = zin.read("[Content_Types].xml").decode("utf-8")

used_rids = {int(m) for m in re.findall(r'Id="rId(\d+)"', rels)}
new_rid = "rId%d" % (max(used_rids) + 1)
used_sheetids = {int(m) for m in re.findall(r'sheetId="(\d+)"', wb)}
new_sheetid = max(used_sheetids) + 1
sheet_part = "xl/worksheets/sheet11.xml"
draw_part = "xl/drawings/drawing6.xml"
assert sheet_part not in names and draw_part not in names, "part already exists"

# ---- 1) drawing: title, button (macro), info panel, note -----------------
EMU_C = 762000    # ~ one of our wide columns
def anchor(c1, r1, c2, r2, body):
    return ('<xdr:twoCellAnchor editAs="oneCell">'
            '<xdr:from><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
            '<xdr:to><xdr:col>%d</xdr:col><xdr:colOff>0</xdr:colOff>'
            '<xdr:row>%d</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
            '%s<xdr:clientData/></xdr:twoCellAnchor>' % (c1, r1, c2, r2, body))

def rpr(sz, b, colour, i=0):
    return ('<a:rPr lang="en-GB" sz="%d" b="%d" i="%d">'
            '<a:solidFill><a:srgbClr val="%s"/></a:solidFill>'
            '<a:latin typeface="Calibri"/></a:rPr>' % (sz, b, i, colour))

def para(runs, align="l", space_after=0):
    return ('<a:p><a:pPr algn="%s"/>%s</a:p>' % (align, runs)) if not space_after else \
           ('<a:p><a:pPr algn="%s"><a:spcAft><a:spcPts val="%d"/></a:spcAft></a:pPr>%s</a:p>'
            % (align, space_after, runs))

def run(text, sz=1100, b=0, colour="3F4448", i=0):
    return rpr(sz, b, colour, i).join(["<a:r>", ""]) + "<a:t>" + text + "</a:t></a:r>"

def textbox(sid, name, x, y, cx, cy, paras, fill=None, line=None, geom="rect",
            macro="", anchor_v="t", ins=None):
    fillxml = ('<a:solidFill><a:srgbClr val="%s"/></a:solidFill>' % fill) if fill else "<a:noFill/>"
    linexml = ('<a:ln w="9525"><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:ln>' % line) if line else '<a:ln><a:noFill/></a:ln>'
    avl = '<a:avLst><a:gd name="adj" fmla="val 12000"/></a:avLst>' if geom == "roundRect" else '<a:avLst/>'
    insxml = ins or 'lIns="91440" tIns="45720" rIns="91440" bIns="45720"'
    return ('<xdr:sp macro="%s" textlink="">'
            '<xdr:nvSpPr><xdr:cNvPr id="%d" name="%s"/>'
            '<xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>'
            '<xdr:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm>'
            '<a:prstGeom prst="%s">%s</a:prstGeom>%s%s</xdr:spPr>'
            '<xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" '
            '%s rtlCol="0" anchor="%s"/><a:lstStyle/>%s</xdr:txBody></xdr:sp>'
            % (macro, sid, name, x, y, cx, cy, geom, avl, fillxml, linexml,
               insxml, anchor_v, paras))

BULLETS = [
    ("Workbook location", "must be a real local folder, not OneDrive / SharePoint"),
    ("Working drive", "the drive holding the \pdf\ working folders"),
    ("SAP GUI session", "open, logged in, scripting enabled"),
    ("PDFCreator", "installed, printer named exactly PDFCreator"),
    ("PDF merger", "GiosPSMC.exe present, or reachable on the network share"),
    ("Cost Centre", "config!B2 is filled in"),
]
bullet_paras = "".join(
    '<a:p><a:pPr algn="l"><a:spcAft><a:spcPts val="360"/></a:spcAft></a:pPr>'
    '<a:r>' + rpr(1050, 1, "1F2225") + '<a:t>' + t + '</a:t></a:r>'
    '<a:r>' + rpr(1050, 0, "6B7176") + '<a:t>  —  ' + d + '</a:t></a:r></a:p>'
    for t, d in BULLETS)

shapes = []
# title
shapes.append(anchor(1, 1, 8, 3, textbox(
    2, "Title", 400000, 200000, 6000000, 500000,
    '<a:p><a:pPr algn="l"/><a:r>' + rpr(2000, 1, "1F2225") +
    '<a:t>Closing Manager — Preflight Check</a:t></a:r></a:p>')))
# subtitle
shapes.append(anchor(1, 3, 8, 5, textbox(
    3, "Subtitle", 400000, 700000, 6000000, 420000,
    '<a:p><a:pPr algn="l"/><a:r>' + rpr(1200, 0, "6B7176") +
    '<a:t>Check that everything this workbook needs is in place '
    '— before you start a month-end close.</a:t></a:r></a:p>')))
# THE BUTTON  -> calls the existing macro, name unchanged
shapes.append(anchor(1, 6, 4, 9, textbox(
    4, "Button Run Preflight Check", 400000, 1500000, 2600000, 700000,
    '<a:p><a:pPr algn="ctr"/><a:r>' + rpr(1400, 1, "FFFFFF") +
    '<a:t>RUN PREFLIGHT CHECK</a:t></a:r></a:p>',
    fill="C7352D", line="9E2A23", geom="roundRect",
    macro="[0]!PreflightCheck", anchor_v="ctr")))
# note beside the button
shapes.append(anchor(4, 6, 9, 9, textbox(
    5, "Note", 3200000, 1500000, 3400000, 700000,
    '<a:p><a:pPr algn="l"/><a:r>' + rpr(1000, 0, "2F6B4F") +
    '<a:t>Safe to run at any time. This check only reads: it starts no close, '
    'creates no folders and writes no files.</a:t></a:r></a:p>',
    anchor_v="ctr")))
# panel heading + bullets
shapes.append(anchor(1, 10, 9, 20, textbox(
    6, "Checks Panel", 400000, 2500000, 6200000, 2600000,
    '<a:p><a:pPr algn="l"><a:spcAft><a:spcPts val="600"/></a:spcAft></a:pPr><a:r>' +
    rpr(1100, 1, "C7352D") + '<a:t>WHAT IT CHECKS</a:t></a:r></a:p>' + bullet_paras,
    fill="F5F6F7", line="DEE1E3")))
# footer note
shapes.append(anchor(1, 21, 9, 23, textbox(
    7, "Footer", 400000, 5300000, 6200000, 500000,
    '<a:p><a:pPr algn="l"/><a:r>' + rpr(900, 0, "93999E") +
    '<a:t>You can also run it from Alt + F8 → PreflightCheck → Run.   '
    '·   Closing Manager V4-CIO   ·   Continuous Improvement Team</a:t></a:r></a:p>')))

drawing = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
           '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
           'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
           + "".join(shapes) + '</xdr:wsDr>')

# ---- 2) the worksheet ----------------------------------------------------
sheet = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
         '<dimension ref="A1"/>'
         '<sheetViews><sheetView showGridLines="0" tabSelected="0" workbookViewId="0"/></sheetViews>'
         '<sheetFormatPr defaultRowHeight="15"/>'
         '<cols><col min="1" max="1" width="3.28515625" customWidth="1"/>'
         '<col min="2" max="10" width="11.7109375" customWidth="1"/></cols>'
         '<sheetData/>'
         '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
         '<drawing r:id="rId1"/></worksheet>')

sheet_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
              '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              '<Relationship Id="rId1" '
              'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" '
              'Target="../drawings/drawing6.xml"/></Relationships>')

# ---- 3) edits to existing parts -----------------------------------------
new_sheet_el = '<sheet name="Preflight" sheetId="%d" r:id="%s"/>' % (new_sheetid, new_rid)
m = re.search(r'<sheet name="START"[^/]*/>', wb)
assert m, "START sheet element not found"
INSERT_AT = 1                      # tab position of the new sheet (0 = START)
wb_new = wb[:m.end()] + new_sheet_el + wb[m.end():]
assert wb_new != wb

# definedName/@localSheetId is a SHEET INDEX, not an id - inserting a tab shifts
# every later sheet, so scoped names must be renumbered or they silently re-point
# at the wrong sheet.
def _bump(mo):
    idx = int(mo.group(1))
    return 'localSheetId="%d"' % (idx + 1 if idx >= INSERT_AT else idx)
wb_new, n_bumped = re.subn(r'localSheetId="(\d+)"', _bump, wb_new)
print("  localSheetId entries renumbered:", n_bumped)

# activeTab is also an index; shift it too if it pointed at or past the insert
def _tab(mo):
    idx = int(mo.group(1))
    return 'activeTab="%d"' % (idx + 1 if idx >= INSERT_AT else idx)
wb_new = re.sub(r'activeTab="(\d+)"', _tab, wb_new)

rels_new = rels.replace('</Relationships>',
    '<Relationship Id="%s" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
    'Target="worksheets/sheet11.xml"/></Relationships>' % new_rid)
assert rels_new != rels

ct_new = ct.replace('</Types>',
    '<Override PartName="/xl/worksheets/sheet11.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/drawings/drawing6.xml" '
    'ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>')
assert ct_new != ct

REPLACE = {
    "xl/workbook.xml": wb_new.encode("utf-8"),
    "xl/_rels/workbook.xml.rels": rels_new.encode("utf-8"),
    "[Content_Types].xml": ct_new.encode("utf-8"),
}
ADD = {
    sheet_part: sheet.encode("utf-8"),
    "xl/worksheets/_rels/sheet11.xml.rels": sheet_rels.encode("utf-8"),
    draw_part: drawing.encode("utf-8"),
}

# ---- 4) write, preserving every other entry verbatim ---------------------
zout = zipfile.ZipFile(OUT, "w")
for item in zin.infolist():
    data = REPLACE.get(item.filename, zin.read(item.filename))
    zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
    zi.compress_type = item.compress_type
    zi.external_attr = item.external_attr
    zi.internal_attr = item.internal_attr
    zi.create_system = item.create_system
    zout.writestr(zi, data)
for name, data in ADD.items():
    zout.writestr(zipfile.ZipInfo(name, date_time=(2026, 7, 29, 12, 0, 0)), data,
                  compress_type=zipfile.ZIP_DEFLATED)
zin.close(); zout.close()
print("wrote %s  (sheet id=%d, %s, macro=[0]!PreflightCheck)" % (OUT, new_sheetid, new_rid))

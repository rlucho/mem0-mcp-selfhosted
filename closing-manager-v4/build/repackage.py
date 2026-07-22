import zipfile, shutil, os
SRC = "workbook.xlsm"
OUT = "Closing_Manager_IP_V4-CIO.xlsm"
newbin = open("vbaProject_v4.bin","rb").read()

zin = zipfile.ZipFile(SRC, "r")
zout = zipfile.ZipFile(OUT, "w")
replaced = False
for item in zin.infolist():
    data = zin.read(item.filename)
    if item.filename == "xl/vbaProject.bin":
        data = newbin; replaced = True
    # preserve original per-entry compression type + metadata
    zi = zipfile.ZipInfo(item.filename, date_time=item.date_time)
    zi.compress_type = item.compress_type
    zi.external_attr = item.external_attr
    zi.internal_attr = item.internal_attr
    zi.create_system = item.create_system
    zout.writestr(zi, data)
zin.close(); zout.close()
assert replaced, "vbaProject.bin not found in original!"
print("wrote", OUT, os.path.getsize(OUT), "bytes; replaced vbaProject.bin =", replaced)

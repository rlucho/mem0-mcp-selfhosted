import struct, sys
import olefile
from oletools.olevba import decompress_stream
from cfb import Cfb, build_cfb
from ovba import compress_ovba, find_source_offset

def _b(x): return x if isinstance(x, bytes) else x.encode("latin-1")

def once(text, old, new, label):
    if text.count(old) != 1:
        sys.exit("FAIL %s: count=%d" % (label, text.count(old)))
    return text.replace(old, new)

def several(text, old, new, n, label):
    if text.count(old) != n:
        sys.exit("FAIL %s: count=%d (expected %d)" % (label, text.count(old), n))
    return text.replace(old, new)

SP_URL_LIT = '"https://troom-x.capgemini.com/sites/InternationalPaper/r2a/_vti_bin/Lists.asmx"'
ARCHIVE_LIT = '"\\\\pl-krabpo-fsc01\\ipa$\\R2R\\R2R - IP EU\\MONTH-END\\CLOSING REPORTS\\"'

# ---- load original bin ----
orig = open("extracted/xl/vbaProject.bin","rb").read()
tpl = Cfb(orig)

# name -> entry index (target names are globally unique)
idx = {}
for i, e in enumerate(tpl.entries):
    nm = tpl.entry_name(e)
    if nm: idx[nm] = i

def module_text(name):
    raw = tpl.entry_data(idx[name])
    off, src_b = find_source_offset(raw)
    return raw, off, src_b.decode("cp1252").replace("\r\n", "\n")

# ---- mCloseEnv content split into declarations (consts) + procedures ----
menv = open("build_v4/mCloseEnv_V4.bas", encoding="utf-8").read().replace("\r\n","\n")
menv = menv.replace("Option Explicit\n", "")
menv = menv.replace(
    "' NEW module. Import it once; it adds the environment / path helpers used by",
    "' V4-CIO environment / path helpers (folded into GlobalModule for the baked build),")
cut = menv.index("'--- TRUE when a path is an http(s) address")
decl_block = menv[:cut].rstrip() + "\n"          # header comments + 3 Public Const
proc_block = menv[cut:].rstrip() + "\n"           # all helper procedures

# ============================ GlobalModule ==================================
graw, goff, gtext = module_text("GlobalModule")

gtext = once(gtext,
 "Global FPath As String, FTemp As String, Fmerger As String, Fmerged As String, Fprinted As String, FFinal As String",
 "Global FPath As String, FTemp As String, Fmerger As String, Fmerged As String, Fprinted As String, FFinal As String\n"
 "Global FShared As String, DiscN As String   'V4-CIO: promoted to module scope so every routine shares one working drive",
 "GM globals")

old_cp = ('Sub CreatePaths()\n\n'
 'Set fso = CreateObject("Scripting.FileSystemObject")\n\n'
 'FPath = ThisWorkbook.Path\n'
 'If Right(FPath, 1) <> "\\" Then FPath = FPath & "\\"\n\n'
 'If fso.DriveExists("D:") Then\n'
 '    DiscN = "D:\\"\n'
 'Else\n'
 '    DiscN = "C:\\"\n'
 'End If\n\n'
 'FTemp = DiscN & "pdf\\temp"\n'
 'Fmerger = DiscN & "pdf\\merger"\n'
 'Fmerged = DiscN & "pdf\\mergedFiles"\n'
 'Fprinted = DiscN & "pdf\\printed"\n'
 'FFinal = DiscN & "pdf\\final"\n'
 'FShared = "\\\\pl-krabpo-fsc01\\ipa$\\R2R\\R2R - IP EU\\MONTH-END\\CLOSING REPORTS\\" & Year(LastDay) & "\\" & Right("0" & Month(LastDay), 2)\n\n'
 'End Sub')

new_cp = ('Sub CreatePaths()\n'
 "'V4-CIO: one consistent working drive (CM_BASE_DRIVE), a local-path guard for\n"
 "'OneDrive/SharePoint, and centralised folder creation. Callers keep using the\n"
 "'same FPath / FTemp / Fmerger / ... globals exactly as before.\n\n"
 'If Not AssertLocalWorkbook() Then End      \'stop hard if the file was opened from the web\n\n'
 'FPath = ThisWorkbook.Path\n'
 'If Right(FPath, 1) <> "\\" Then FPath = FPath & "\\"\n\n'
 'DiscN = CM_BASE_DRIVE                       \'V4-CIO FIX: was "D:\\ if present else C:\\"; now one value everywhere\n'
 'FTemp = DiscN & "pdf\\temp"\n'
 'Fmerger = DiscN & "pdf\\merger"\n'
 'Fmerged = DiscN & "pdf\\mergedFiles"\n'
 'Fprinted = DiscN & "pdf\\printed"\n'
 'FFinal = DiscN & "pdf\\final"\n'
 'FShared = "\\\\pl-krabpo-fsc01\\ipa$\\R2R\\R2R - IP EU\\MONTH-END\\CLOSING REPORTS\\" & Year(LastDay) & "\\" & Right("0" & Month(LastDay), 2)\n\n'
 'Call EnsureFolders                          \'build \\pdf\\* + report tree, provision GiosPSMC.exe\n\n'
 'End Sub')

gtext = once(gtext, old_cp, new_cp, "GM CreatePaths")
# endpoint centralisation (BEFORE folding in the const defs, so the const
# definitions themselves keep the string literals):
gtext = once(gtext, "FShared = " + ARCHIVE_LIT, "FShared = CM_ARCHIVE_ROOT", "GM FShared const")
gtext = several(gtext, "Url = " + SP_URL_LIT, "Url = CM_SP_BASE", 4, "GM SharePoint URL const")
# insert const declarations before the first procedure
gtext = once(gtext, "\nSub CreatePaths()", "\n" + decl_block + "\nSub CreatePaths()", "GM insert decls")
# append helper procedures at the end
gtext = gtext.rstrip("\n") + "\n\n" + proc_block

# ============================== Closing =====================================
craw, coff, ctext = module_text("Closing")
old_block = ('DiscN = "C:\\"\n\n'
 'If fso.FolderExists(DiscN & "pdf\\") = False Then fso.CreateFolder (DiscN & "pdf\\")\n'
 'If fso.FolderExists(DiscN & "pdf\\mergedFiles\\") = False Then fso.CreateFolder (DiscN & "pdf\\mergedFiles\\")\n'
 'If fso.FolderExists(DiscN & "pdf\\merger\\") = False Then fso.CreateFolder (DiscN & "pdf\\merger\\")\n'
 'If fso.FolderExists(DiscN & "pdf\\temp\\") = False Then fso.CreateFolder (DiscN & "pdf\\temp\\")\n'
 'If fso.FolderExists(DiscN & "pdf\\printed\\") = False Then fso.CreateFolder (DiscN & "pdf\\printed\\")\n'
 'If fso.FolderExists(DiscN & "pdf\\final\\") = False Then fso.CreateFolder (DiscN & "pdf\\final\\")\n'
 'If fso.FileExists(DiscN & "pdf\\merger\\GiosPSMC.exe") = False Then fso.CopyFile "\\\\pl-krabpo-fsc01\\ipa$\\R2R\\R2R - IP EU GL West\\USEFUL\\pdf\\merger\\GiosPSMC.exe", DiscN & "pdf\\merger\\GiosPSMC.exe"\n'
 'If fso.FolderExists("C:\\_Files to Transfer\\MONTH END CLOSE\\" & Year(LastDay) & "\\") = False Then fso.CreateFolder ("C:\\_Files to Transfer\\MONTH END CLOSE\\" & Year(LastDay) & "\\")\n'
 'If fso.FolderExists("C:\\_Files to Transfer\\MONTH END CLOSE\\" & Year(LastDay) & "\\" & Right("0" & Month(LastDay), 2) & "\\") = False Then fso.CreateFolder ("C:\\_Files to Transfer\\MONTH END CLOSE\\" & Year(LastDay) & "\\" & Right("0" & Month(LastDay), 2) & "\\")')
new_block = 'Call EnsureFolders   \'V4-CIO FIX: single-drive, parent-aware folder creation (was hardcoded C:\\ here while CreatePaths used D:\\ if present -> broke print/merge on D:-drive PCs).'
ctext = once(ctext, old_block, new_block, "Closing folder block")

# ============================== Printing ====================================
praw, poff, ptext = module_text("Printing")
old_pdf = ('If fso.FileExists("C:\\Program Files\\PDFCreator\\PDFCreator.exe") Then\n'
 '    varProc = Shell("C:\\Program Files\\PDFCreator\\PDFCreator.exe", 1)\n'
 'Else\n'
 '    varProc = Shell("C:\\Program Files (x86)\\PDFCreator\\PDFCreator.exe", 1)\n'
 'End If')
new_pdf = ('If fso.FileExists("C:\\Program Files\\PDFCreator\\PDFCreator.exe") Then\n'
 '    varProc = Shell("C:\\Program Files\\PDFCreator\\PDFCreator.exe", 1)\n'
 'ElseIf fso.FileExists("C:\\Program Files (x86)\\PDFCreator\\PDFCreator.exe") Then\n'
 '    varProc = Shell("C:\\Program Files (x86)\\PDFCreator\\PDFCreator.exe", 1)\n'
 'Else\n'
 "    'V4-CIO FIX: was an unguarded Shell of the (x86) path -> run-time error if PDFCreator absent\n"
 '    MsgBox "PDFCreator is not installed in the expected location." & vbCrLf & _\n'
 '           "Install PDFCreator (the printer must be named \'PDFCreator\') before running the close.", _\n'
 '           vbCritical, "Closing Manager"\n'
 'End If')
ptext = once(ptext, old_pdf, new_pdf, "SetPDFCreator guard")

old_combine = ('Sub CombinePDF(printN)\n\n'
 'Set shellX = CreateObject("WScript.Shell")\n'
 'Set fso = CreateObject("scripting.filesystemobject")\n\n'
 'Call CreateVariants\n'
 'Call CreatePaths\n\n'
 'paramsource = ""\n'
 'paramOutput = Fmerged & "\\" & Yearx & Monthx & CC & ".pdf"\n'
 '       \n'
 'For i = 1 To printN\n'
 '    paramsource = paramsource & Chr(34) & FFinal & "\\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & i & ".pdf" & Chr(34) & " "\n'
 'Next i\n'
 '          \n'
 'paramsource = Left(paramsource, Len(paramsource) - 1)\n'
 'shellX.Run "%COMSPEC% /c " & Fmerger & "\\GiosPSMC.exe" & " " & paramsource & " output " & paramOutput\n\n'
 'Do\n'
 '    If fso.FileExists(paramOutput) Then\n'
 '        File = fso.GetFile(paramOutput)\n'
 '        Exit Do\n'
 '    Else\n'
 '        Application.Wait (Now + TimeValue("0:00:01"))\n'
 '    End If\n'
 'Loop\n'
 'Do\n'
 '    size1 = fso.GetFile(File).Size\n'
 '    Application.Wait (Now + TimeValue("0:00:01"))\n'
 '    size2 = fso.GetFile(File).Size\n\n'
 '    If size1 = size2 And size1 <> 0 And size2 <> 0 Then\n'
 '        Exit Do\n'
 '    End If\n'
 'Loop\n\n'
 'fso.DeleteFile (FFinal & "\\*.*"), True\n\n'
 'fN = 1\n'
 'FPathReport = "C:\\_Files to Transfer\\MONTH END CLOSE\\" & Yearx & "\\" & Right("0" & Monthx, 2) & "\\"\n'
 'FName = CC & " Closure " & Right("0" & Monthx, 2) & " " & Yearx & ".pdf"\n'
 'FName1 = CC & " Closure " & Right("0" & Monthx, 2) & " " & Yearx\n'
 'Do Until fso.FileExists(FPathReport & FName) = False\n'
 '    FName = FName1 & "-" & fN & ".pdf"\n'
 'Loop\n\n'
 'fso.MoveFile File, FPathReport & FName\n\n'
 'End Sub')

new_combine = ('Sub CombinePDF(printN)\n'
 "'V4-CIO: same merge behaviour, but the waits are now bounded (cannot hang Excel),\n"
 "'the merge tool is checked before use, and the destination-name loop no longer\n"
 "'spins forever. Functional logic and the GiosPSMC command line are unchanged.\n\n"
 'Dim waited As Long\n'
 "Const MAX_WAIT As Long = 300     'safety cap in seconds for each wait loop\n\n"
 'Set shellX = CreateObject("WScript.Shell")\n'
 'Set fso = CreateObject("scripting.filesystemobject")\n\n'
 'Call CreateVariants\n'
 'Call CreatePaths\n\n'
 'paramsource = ""\n'
 'paramOutput = Fmerged & "\\" & Yearx & Monthx & CC & ".pdf"\n'
 '       \n'
 'For i = 1 To printN\n'
 '    paramsource = paramsource & Chr(34) & FFinal & "\\" & CC & "-" & Yearx & "-" & Right("0" & Monthx, 2) & "-" & i & ".pdf" & Chr(34) & " "\n'
 'Next i\n'
 '          \n'
 'paramsource = Left(paramsource, Len(paramsource) - 1)\n\n'
 "'V4-CIO FIX: fail fast with a clear message if the merge tool is missing\n"
 'If Not fso.FileExists(Fmerger & "\\GiosPSMC.exe") Then\n'
 '    MsgBox "PDF merger not found:" & vbCrLf & Fmerger & "\\GiosPSMC.exe" & vbCrLf & vbCrLf & _\n'
 '           "Run \'Preflight Check\' or restore it from the network share, then try again.", _\n'
 '           vbCritical, "Closing Manager"\n'
 '    Exit Sub\n'
 'End If\n\n'
 'shellX.Run "%COMSPEC% /c " & Fmerger & "\\GiosPSMC.exe" & " " & paramsource & " output " & paramOutput\n\n'
 "'V4-CIO FIX: bounded wait for the merged file to appear (was an unbounded Do..Loop)\n"
 'waited = 0\n'
 'Do\n'
 '    If fso.FileExists(paramOutput) Then\n'
 '        File = fso.GetFile(paramOutput)\n'
 '        Exit Do\n'
 '    Else\n'
 '        Application.Wait (Now + TimeValue("0:00:01"))\n'
 '        waited = waited + 1\n'
 '    End If\n'
 '    If waited >= MAX_WAIT Then\n'
 '        MsgBox "PDF merge timed out - the merged file was not created:" & vbCrLf & paramOutput, _\n'
 '               vbCritical, "Closing Manager"\n'
 '        Exit Sub\n'
 '    End If\n'
 'Loop\n\n'
 "'V4-CIO FIX: bounded wait for the file to stop growing (was an unbounded Do..Loop)\n"
 'Do\n'
 '    size1 = fso.GetFile(File).Size\n'
 '    Application.Wait (Now + TimeValue("0:00:01"))\n'
 '    size2 = fso.GetFile(File).Size\n'
 '    waited = waited + 1\n\n'
 '    If size1 = size2 And size1 <> 0 And size2 <> 0 Then\n'
 '        Exit Do\n'
 '    End If\n'
 '    If waited >= MAX_WAIT * 2 Then Exit Do\n'
 'Loop\n\n'
 'fso.DeleteFile (FFinal & "\\*.*"), True\n\n'
 "Call EnsureFolders   'V4-CIO: make sure the year\\month tree exists before the move\n\n"
 'fN = 1\n'
 'FPathReport = CM_REPORT_ROOT & Yearx & "\\" & Right("0" & Monthx, 2) & "\\"\n'
 'FName = CC & " Closure " & Right("0" & Monthx, 2) & " " & Yearx & ".pdf"\n'
 'FName1 = CC & " Closure " & Right("0" & Monthx, 2) & " " & Yearx\n'
 'Do Until fso.FileExists(FPathReport & FName) = False\n'
 '    FName = FName1 & "-" & fN & ".pdf"\n'
 '    fN = fN + 1                          \'V4-CIO FIX: was missing -> infinite loop when a same-named report already existed\n'
 'Loop\n\n'
 "'V4-CIO FIX: report, don't crash, if the final move fails\n"
 'On Error Resume Next\n'
 'fso.MoveFile File, FPathReport & FName\n'
 'If Err.Number <> 0 Then\n'
 '    MsgBox "Could not move the final report to:" & vbCrLf & FPathReport & FName & vbCrLf & vbCrLf & _\n'
 '           Err.Description, vbCritical, "Closing Manager"\n'
 'End If\n'
 'On Error GoTo 0\n\n'
 'End Sub')
ptext = once(ptext, old_combine, new_combine, "CombinePDF")

# ============================== Admin =======================================
# Endpoint centralisation only: point the active SharePoint calls at CM_SP_BASE
# (defined as a Public Const in GlobalModule). 3 literals: 2 active + 1 comment.
araw, aoff, atext = module_text("Admin")
atext = several(atext, "Url = " + SP_URL_LIT, "Url = CM_SP_BASE", 3, "Admin SharePoint URL const")

# ---- assemble new module streams (keep p-code prefix, replace source) ----
def new_stream(raw, off, text):
    src = text.replace("\n", "\r\n").encode("cp1252")
    return raw[:off] + compress_ovba(src)

# ---- V4-CIO post-pass: bounded print waits, live progress, plain-language errors ----
import post_v4
praw2, poff2, potext = module_text("Postings")
potext = post_v4.postings(potext)
gtext = post_v4.review_fixes_globalmodule(gtext)
atext = post_v4.admin(atext)
ptext = post_v4.printing(ptext)
ctext = post_v4.closing(ctext)
print("post-pass: 10 print waits bounded, %d breadcrumbs, CM_Fail handler" % len(post_v4.STEPS))

new_data = {
    idx["GlobalModule"]: new_stream(graw, goff, gtext),
    idx["Closing"]:      new_stream(craw, coff, ctext),
    idx["Printing"]:     new_stream(praw, poff, ptext),
    idx["Admin"]:        new_stream(araw, aoff, atext),
    idx["Postings"]:     new_stream(praw2, poff2, potext),
}

# ---- force clean recompile: bump _VBA_PROJECT version, empty __SRP_* caches ----
vp = bytearray(tpl.entry_data(idx["_VBA_PROJECT"]))
vp[2:4] = b"\xFF\xFF"
new_data[idx["_VBA_PROJECT"]] = bytes(vp)

srp = [nm for nm in idx if nm.startswith("__SRP_")]
for nm in srp:
    new_data[idx[nm]] = b""

# ---- build final bin ----
final = build_cfb(tpl, new_data)
open("vbaProject_v4.bin","wb").write(final)

# save the V4 module sources (exactly what is baked into the workbook)
for nm, txt in (("GlobalModule_FOLDED", gtext), ("Admin", atext), ("Closing", ctext), ("Printing", ptext), ("Postings", potext)):
    open("build_v4/%s.bas" % nm, "w", encoding="utf-8").write(txt.replace("\n", "\r\n"))

print("built vbaProject_v4.bin  size=%d (orig=%d)" % (len(final), len(orig)))
print("SRP streams emptied:", len(srp))
print("OK")

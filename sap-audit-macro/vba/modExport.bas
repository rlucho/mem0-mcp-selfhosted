Attribute VB_Name = "modExport"
'=======================================================================
' modExport -- gets a SAP list out to a local file.
'
' Two routes, because SAP has two kinds of list:
'   ExportAlvGrid   for the ALV controls FEBAN and FBL1N use
'   ExportClassicList for the older non-ALV lists, via the %pc OK-code
'
' Both are read-only on the SAP side. The only thing written is a file on
' your own filesystem.
'
' Note on format: 'Text with tabs' is preferred over 'Spreadsheet'. The
' spreadsheet route can hand the list to a live Excel instance instead of
' a file, which then fights the workbook this macro is running from.
'=======================================================================
Option Explicit

'-----------------------------------------------------------------------
' Where SAP is told to write, and where the evidence ends up.
'
' SAP does not merely write the export -- it opens it in Excel afterwards.
' It does that through OLE automation, driving this very instance, so
' Excel's 'ignore other applications that use DDE' switch never sees the
' request and cannot refuse it. Excel will not hold two workbooks with the
' same name whatever folders they are in, and since each sample got its own
' folder every sample writes the same handful of names. Sample 42's
' '4 - Documents behind the largest payment.xlsx' therefore arrives while
' sample 41's is still open, and a modal stops the run until a human clicks
' OK. Closing the previous copy first cannot win that race: the hand-off
' lands whenever SAP gets round to it, seconds later, and three separate
' closes did not stop it.
'
' So there is no race here any more. SAP writes to a scratch folder under a
' name carrying a counter that never repeats, and the file is copied into
' the sample folder under its proper name once it is completely written.
' Excel can open the scratch copy whenever it likes -- no two of those are
' ever called the same thing -- and the sample folder only ever receives a
' file that nothing has open.
'
' Timing is the whole trick. Between samples the scratch WORKBOOKS are
' closed but the FILES are left alone, because SAP's hand-off can still be
' in flight and Excel asked to open a file that has just been deleted stops
' the run just as dead as the duplicate name did. Deleting happens at the
' start of a run, when nothing can be pending, and at the end with a grace
' period for the last export.
'
' It sits under TEMP, not under the evidence root, so an auditor opening the
' pack never sees it at all.
'-----------------------------------------------------------------------
Private Const HANDOVER_FOLDER As String = "sap-audit-handover"

' How long a scratch file is left alone at the end of a run, in case SAP is
' still on its way to Excel with it. Generous on purpose -- the cost of
' waiting is a few files sitting in TEMP until the next run opens, and the
' cost of not waiting is a modal that stops the run.
Public Const HANDOVER_GRACE_SECONDS As Double = 120

Private mHandoverSequence As Long

'-----------------------------------------------------------------------
' ALV grid export via the grid's own toolbar export button.
' Returns the full path written, or "" if nothing was written.
'
' recordings/Audit.vbs never exported the FEBAN result grid -- it only
' exported the classic cleared-items list -- so these three keys are not
' recorded. That makes this whole step optional: when they are blank it
' logs and returns rather than raising, because a missing nice-to-have
' export must not take down the sample it belongs to.
'-----------------------------------------------------------------------
Public Function ExportAlvGrid(ByVal sampleIdx As Long, ByVal folder As String, _
                              ByVal fileName As String) As String
    ExportAlvGrid = ExportAlvGridById(sampleIdx, "FEBAN.ResultGrid", folder, fileName)
End Function

' Same, for any mapped ALV grid -- the FEBAN statement list and the FBL1N
' payment list both come out this way.
Public Function ExportAlvGridById(ByVal sampleIdx As Long, ByVal gridKey As String, _
                                  ByVal folder As String, ByVal fileName As String) As String
    Dim grid As Object
    Dim target As String
    Dim gridId As String
    Dim toolbarButton As String, menuItem As String

    gridId = modConfig.ElementIdOrBlank(gridKey)
    toolbarButton = modConfig.ElementIdOrBlank("Export.AlvToolbarButton")
    menuItem = modConfig.ElementIdOrBlank("Export.AlvMenuItem")

    If Len(gridId) = 0 Or Len(toolbarButton) = 0 Or Len(menuItem) = 0 Then
        modLog.LogAction sampleIdx, "Export list", _
                     "Skipped -- " & gridKey & ", Export.AlvToolbarButton or " & _
                     "Export.AlvMenuItem is blank on the Screen Map.", _
                     "SKIPPED", vbNullString
        Exit Function
    End If

    If Not modSapConnect.Exists(gridId) Then
        modLog.LogAction sampleIdx, "Export list", _
                     "Skipped -- " & gridKey & " is not on the current screen.", _
                     "SKIPPED", vbNullString
        Exit Function
    End If

    target = modUtil.JoinPath(folder, fileName)
    If modSafety.BlockedByDryRun("Would export " & gridKey & " to " & target) Then Exit Function

    modUtil.EnsureFolder folder
    Set grid = modSapConnect.Element(gridId)

    ' The sequence Audit2.vbs captured: &MB_EXPORT then &XXL.
    On Error Resume Next
    grid.pressToolbarContextButton toolbarButton
    grid.selectContextMenuItem menuItem
    On Error GoTo 0
    modSapConnect.WaitForSap

    If Not modSapConnect.ModalWindowOpen() Then
        modLog.LogAction sampleIdx, "Export list", _
                     "The export menu produced no dialog. Check Export.AlvToolbarButton " & _
                     "and Export.AlvMenuItem against a recording of that right-click.", _
                     "SKIPPED", vbNullString
        Exit Function
    End If

    ConfirmFormatPopupIfPresent
    ExportAlvGridById = CompleteSaveDialog(sampleIdx, folder, fileName)
End Function

'-----------------------------------------------------------------------
' Classic (non-ALV) list export, via the menu path the recording used:
' List > Save/Send > File. Both are display-side functions.
'
' Export.ListMenu is preferred over the %pc OK-code because that is what
' recordings/Audit.vbs actually did on this system, and %pc is not offered
' on every classic list.
'-----------------------------------------------------------------------
Public Function ExportClassicList(ByVal sampleIdx As Long, ByVal folder As String, _
                                  ByVal fileName As String) As String
    ExportClassicList = ExportListFrom(sampleIdx, "wnd[0]", vbNullString, folder, fileName)
End Function

' Export whichever list is showing, from whichever window holds it.
'
' The Payment Usage list is in wnd[0] in the recordings and in a modal on
' this system, and the two need different routes: a modal has no menu bar,
' so List > Save/Send > File does not exist there, but the list inside it is
' an ALV grid and exports through the grid's own toolbar. Work out which
' shape is in front of us rather than assuming either.
Public Function ExportListFrom(ByVal sampleIdx As Long, ByVal windowId As String, _
                               ByVal gridId As String, ByVal folder As String, _
                               ByVal fileName As String) As String
    Dim target As String
    Dim menuId As String

    target = modUtil.JoinPath(folder, fileName)

    If modSafety.BlockedByDryRun("Would export the list to " & target) Then Exit Function

    modUtil.EnsureFolder folder

    ' Order matters. List > Save/Send > File is what the recordings do and
    ' what works by hand, so try that first and only fall back to the grid's
    ' own toolbar when the menu is genuinely not there -- which is the case
    ' when the list sits inside a modal, since a modal has no menu bar.
    menuId = modConfig.ElementIdOrBlank("Export.ListMenu")

    If Len(menuId) > 0 And modSapConnect.Exists(menuId) Then
        modLog.LogAction sampleIdx, "Export list", _
                     "Exporting through List > Save/Send > File.", "OK", vbNullString
        modSapConnect.Element(menuId).Select
        modSapConnect.WaitForSap
    Else
        If Len(gridId) = 0 Then gridId = FirstGridIn(windowId)

        If Len(gridId) > 0 Then
            modLog.LogAction sampleIdx, "Export list", _
                         "No list menu available, so exporting the grid " & gridId & _
                         " in " & windowId & " through the ALV toolbar instead.", _
                         "OK", vbNullString

            On Error Resume Next
            modSapConnect.Element(gridId).pressToolbarContextButton _
                modConfig.ElementIdOrBlank("Export.AlvToolbarButton")
            modSapConnect.Element(gridId).selectContextMenuItem _
                modConfig.ElementIdOrBlank("Export.AlvMenuItem")
            On Error GoTo 0
            modSapConnect.WaitForSap
        Else
            modLog.LogAction sampleIdx, "Export list", _
                         "Neither a list menu nor a grid was found in " & windowId & _
                         ". Falling back to the %pc OK-code.", "MANUAL", vbNullString
            modSafety.GuardedOkCode "%pc"
        End If
    End If

    ConfirmFormatPopupIfPresent
    ExportListFrom = CompleteSaveDialogAnyWindow(sampleIdx, folder, fileName)
End Function

' The first grid-like control in a window, or "" when there is none.
Public Function FirstGridIn(ByVal windowId As String) As String
    FirstGridIn = SearchForGrid(windowId, 0)
End Function

Private Function SearchForGrid(ByVal elementId As String, ByVal depth As Long) As String
    Dim control As Object, child As Object
    Dim kind As String, found As String

    If depth > 10 Then Exit Function
    If Not modSapConnect.Exists(elementId) Then Exit Function

    ' The title bar carries a shell of its own, and it answers RowCount, so it
    ' passes for a grid. Sample 141 exported wnd[0]/titl/shellcont[1]/shell --
    ' the title bar -- and reported an export failure, when what had actually
    ' gone wrong was three steps earlier and nothing was on screen to export.
    If InStr(1, elementId, "/titl/", vbTextCompare) > 0 Then Exit Function

    Set control = modSapConnect.Element(elementId)

    On Error Resume Next
    kind = control.Type
    On Error GoTo 0

    If kind = "GuiGridView" Or kind = "GuiShell" Then
        ' A shell that answers RowCount is a grid; a toolbar shell is not.
        On Error Resume Next
        If control.RowCount >= 0 Then SearchForGrid = control.Id
        On Error GoTo 0
        If Len(SearchForGrid) > 0 Then Exit Function
    End If

    On Error Resume Next
    For Each child In control.Children
        found = SearchForGrid(child.Id, depth + 1)
        If Len(found) > 0 Then
            SearchForGrid = found
            Exit For
        End If
    Next child
    On Error GoTo 0
End Function

' The save dialog can come up in wnd[1] or, when the list itself is already
' modal, in wnd[2]. Take whichever one is carrying the path field.
Private Function CompleteSaveDialogAnyWindow(ByVal sampleIdx As Long, ByVal folder As String, _
                                             ByVal fileName As String) As String
    If modSapConnect.Exists("wnd[2]/usr/ctxtDY_PATH") Then
        CompleteSaveDialogAnyWindow = CompleteSaveDialogIn(sampleIdx, "wnd[2]", folder, fileName)
    ElseIf modSapConnect.Exists("wnd[1]/usr/ctxtDY_PATH") Then
        CompleteSaveDialogAnyWindow = CompleteSaveDialogIn(sampleIdx, "wnd[1]", folder, fileName)
    Else
        modLog.LogAction sampleIdx, "Export list", _
                     "No save dialog appeared in wnd[1] or wnd[2] after the export " & _
                     "command, so nothing was written.", "ERROR", vbNullString
    End If
End Function

' Same save dialog, for callers that opened it themselves -- the invoice
' download in modChain reaches it from the attachment list, not from a list
' export.
Public Function CompleteSaveDialogPublic(ByVal sampleIdx As Long, ByVal folder As String, _
                                         ByVal fileName As String) As String
    modUtil.EnsureFolder folder
    CompleteSaveDialogPublic = CompleteSaveDialog(sampleIdx, folder, fileName)
End Function

' Same, for a save dialog that opened somewhere other than wnd[1]. The
' attachment export lands in wnd[2], because the attachment list itself is
' already occupying wnd[1].
Public Function CompleteSaveDialogIn(ByVal sampleIdx As Long, ByVal windowId As String, _
                                     ByVal folder As String, ByVal fileName As String) As String
    Dim pathId As String, nameId As String
    Dim writeFolder As String, writeName As String
    Dim written As String, target As String
    Dim bytes As Double

    modUtil.EnsureFolder folder

    ' Overwriting silently would destroy evidence from an earlier run.
    If modUtil.FileExists(modUtil.JoinPath(folder, fileName)) Then
        fileName = UniqueName(folder, fileName)
    End If

    ' Rebase the recorded wnd[1] paths onto whichever window this dialog is in,
    ' so one set of Save.* rows serves both depths.
    pathId = Rebase(modConfig.ElementId("Save.Path"), windowId)
    nameId = Rebase(modConfig.ElementId("Save.FileName"), windowId)

    If Not modSapConnect.Exists(pathId) Or Not modSapConnect.Exists(nameId) Then
        modLog.LogAction sampleIdx, "Export", _
                     "The save dialog in " & windowId & " does not carry " & pathId & _
                     " / " & nameId & ", so the file could not be placed.", _
                     "ERROR", vbNullString
        Exit Function
    End If

    HandoverTarget folder, fileName, writeFolder, writeName

    modSapConnect.Element(pathId).Text = writeFolder
    modSapConnect.Element(nameId).Text = writeName

    modSafety.GuardedPress Rebase(modConfig.ElementId("Save.GenerateButton"), windowId)
    modSapConnect.WaitForSap

    written = modUtil.JoinPath(writeFolder, writeName)
    If Not WaitForFile(written, 20) Then
        modLog.LogAction sampleIdx, "Export", _
                     "SAP reported no error but " & written & " did not appear.", _
                     "ERROR", written
        Exit Function
    End If

    bytes = WaitUntilStable(written, 20)

    target = Deliver(sampleIdx, written, folder, fileName)
    If Len(target) = 0 Then Exit Function

    modLog.LogAction sampleIdx, "Export", _
                 "Wrote " & Format$(bytes / 1024, "0.0") & " KB", "OK", target

    CompleteSaveDialogIn = target
End Function

' 'wnd[1]/usr/ctxtDY_PATH' + 'wnd[2]' -> 'wnd[2]/usr/ctxtDY_PATH'
Private Function Rebase(ByVal elementId As String, ByVal windowId As String) As String
    Dim slash As Long

    slash = InStr(elementId, "/")
    If slash = 0 Then
        Rebase = elementId
    Else
        Rebase = windowId & Mid$(elementId, slash)
    End If
End Function

'-----------------------------------------------------------------------
' The 'in which format?' popup. Not every release shows it, and not every
' list offers the same options, so this is best-effort and never fatal.
'-----------------------------------------------------------------------
Private Sub ConfirmFormatPopupIfPresent()
    Dim radioId As String, okId As String

    If Not modSapConnect.ModalWindowOpen() Then Exit Sub

    modSafety.AssertPopupKnown

    radioId = modConfig.ElementIdOrBlank("Export.AlvFormatRadio")
    If Len(radioId) > 0 Then
        If modSapConnect.Exists(radioId) Then
            On Error Resume Next
            modSapConnect.Element(radioId).Select
            On Error GoTo 0
        End If
    End If

    okId = modConfig.ElementIdOrBlank("Export.FormatOkButton")
    If Len(okId) > 0 Then
        If modSapConnect.Exists(okId) Then modSafety.GuardedPress okId
    Else
        modSafety.GuardedSendVKey "wnd[1]", 0
    End If

    modSapConnect.WaitForSap
End Sub

'-----------------------------------------------------------------------
' The save-as dialog: fill in path and name, press Generate, then verify
' the file is actually on disk before claiming success.
'-----------------------------------------------------------------------
Private Function CompleteSaveDialog(ByVal sampleIdx As Long, ByVal folder As String, _
                                    ByVal fileName As String) As String
    Dim pathId As String, nameId As String, encodingId As String
    Dim writeFolder As String, writeName As String
    Dim written As String, target As String
    Dim bytes As Double

    If Not modSapConnect.ModalWindowOpen() Then
        modLog.LogAction sampleIdx, "Export", _
                     "No save dialog appeared. The list may be empty, or the export " & _
                     "menu item may be wrong for this release.", "ERROR", vbNullString
        Exit Function
    End If

    modSafety.AssertPopupKnown

    pathId = modConfig.ElementId("Save.Path")
    nameId = modConfig.ElementId("Save.FileName")

    ' Overwriting silently would destroy evidence from an earlier run.
    If modUtil.FileExists(modUtil.JoinPath(folder, fileName)) Then
        fileName = UniqueName(folder, fileName)
    End If

    ' The recording left the default file name in place, so this field's ID is
    ' standard rather than confirmed. If it is not on the dialog, SAP keeps its
    ' own default name -- which still exports, just not where the caller
    ' expects, so that has to be an error rather than a shrug.
    If Not modSapConnect.Exists(nameId) Then
        modLog.LogAction sampleIdx, "Export", _
                     "The save dialog has no field " & nameId & ", so the file name " & _
                     "could not be set and SAP would write its own default. Record " & _
                     "this dialog and correct Save.FileName on the Screen Map.", _
                     "ERROR", vbNullString
        Exit Function
    End If

    HandoverTarget folder, fileName, writeFolder, writeName

    modSapConnect.Element(pathId).Text = writeFolder
    modSapConnect.Element(nameId).Text = writeName

    encodingId = modConfig.ElementIdOrBlank("Save.Encoding")
    If Len(encodingId) > 0 Then
        If modSapConnect.Exists(encodingId) Then
            On Error Resume Next
            modSapConnect.Element(encodingId).Text = "4110"    ' UTF-8
            On Error GoTo 0
        End If
    End If

    modSafety.GuardedPress modConfig.ElementId("Save.GenerateButton")
    modSapConnect.WaitForSap

    ' SAP writes asynchronously; give the file a moment to land.
    written = modUtil.JoinPath(writeFolder, writeName)
    If Not WaitForFile(written, 15) Then
        modLog.LogAction sampleIdx, "Export", _
                     "SAP reported no error but " & written & " did not appear.", _
                     "ERROR", written
        Exit Function
    End If

    bytes = WaitUntilStable(written, 20)
    If bytes = 0 Then
        modLog.LogAction sampleIdx, "Export", _
                     "Wrote " & written & " but it is empty -- the list had no rows.", _
                     "ERROR", written
        Exit Function
    End If

    target = Deliver(sampleIdx, written, folder, fileName)
    If Len(target) = 0 Then Exit Function

    modLog.LogAction sampleIdx, "Export", _
                 "Wrote " & Format$(bytes / 1024, "0.0") & " KB", "OK", target

    CompleteSaveDialog = target
End Function

'-----------------------------------------------------------------------
' The scratch hand-off. See the note at the top of the module.
'-----------------------------------------------------------------------

' Decide where SAP should be told to put this file. Normally that is the
' scratch folder under a name that never repeats; if the scratch folder is
' unavailable it is the sample folder itself, which is what this did before
' and still works, just with the collision back.
Private Sub HandoverTarget(ByVal folder As String, ByVal fileName As String, _
                           ByRef writeFolder As String, ByRef writeName As String)
    Dim scratch As String

    scratch = HandoverFolder()

    If Len(scratch) = 0 Or Not NeedsHandover(fileName) Then
        writeFolder = folder
        writeName = fileName
        Exit Sub
    End If

    On Error GoTo NoScratch
    modUtil.EnsureFolder scratch

    ' A leftover from a run that was killed mid-export would satisfy the
    ' 'has it appeared yet' wait without SAP writing anything, so step over
    ' any name already taken rather than reusing it.
    Do
        mHandoverSequence = mHandoverSequence + 1
        writeName = Format$(mHandoverSequence, "0000") & " " & fileName
    Loop While modUtil.FileExists(modUtil.JoinPath(scratch, writeName))

    writeFolder = scratch
    Exit Sub

NoScratch:
    writeFolder = folder
    writeName = fileName
End Sub

' TEMP, because nothing temporary belongs in a pack an auditor receives.
Private Function HandoverFolder() As String
    Dim base As String

    base = Environ$("TEMP")
    If Len(base) = 0 Then base = Environ$("TMP")

    ' No TEMP at all is unusual but not fatal, and the evidence root is known
    ' to be writable -- it is where the run is already putting files. The
    ' folder gets swept either way.
    If Len(base) = 0 Then base = modConfig.DownloadRoot()
    If Len(base) = 0 Then Exit Function

    HandoverFolder = modUtil.JoinPath(base, HANDOVER_FOLDER)
End Function

' Only spreadsheets need it. Excel is the only thing that objects to a
' repeated name, and the invoice PDF is the only other kind of file the run
' writes -- routing that through the scratch folder would just risk leaving
' a copy behind when Acrobat holds it open.
Private Function NeedsHandover(ByVal fileName As String) As Boolean
    Dim dotPos As Long
    Dim extension As String

    dotPos = InStrRev(fileName, ".")
    If dotPos = 0 Then Exit Function

    extension = LCase$(Mid$(fileName, dotPos))
    NeedsHandover = (extension = ".xlsx" Or extension = ".xls" Or _
                     extension = ".csv" Or extension = ".txt")
End Function

' Put the finished file where the pack expects it, and answer with where
' that is. A copy rather than a move: SAP's hand-off to Excel can still be
' on its way, and 'the file could not be found' is exactly the modal this
' is here to avoid. The scratch copy goes in the next sweep.
Private Function Deliver(ByVal sampleIdx As Long, ByVal written As String, _
                         ByVal folder As String, ByVal fileName As String) As String
    Dim target As String

    target = modUtil.JoinPath(folder, fileName)

    ' Written in place already -- no scratch folder was available.
    If StrComp(written, target, vbTextCompare) = 0 Then
        Deliver = target
        Exit Function
    End If

    modUtil.EnsureFolder folder

    ' Defensive. Nothing should be holding this name: this line is the only
    ' thing that ever creates it, and the macro closes what it opens.
    modUtil.CloseWorkbooksNamed fileName

    If Not modUtil.CopyFileTo(written, target) Then
        modLog.LogAction sampleIdx, "Export", _
                     "SAP wrote " & written & " but it could not be copied to " & _
                     target & ", so the sample folder has no copy of it.", _
                     "ERROR", written
        Exit Function
    End If

    Deliver = target
End Function

'-----------------------------------------------------------------------
' Close whatever Excel has opened out of the scratch folder, WITHOUT
' deleting anything.
'
' This is the between-samples call, and the distinction matters. SAP hands
' each export to Excel on its own schedule -- seconds after the file is
' written, well into the next sample -- and if the file has been deleted by
' then Excel puts up 'Sorry, we couldn't find ...' and waits for a human.
' That is the same class of stoppage the scratch folder exists to prevent,
' so during a run the files stay put and only the workbooks get closed.
' Closing is enough: it is the open workbooks that would otherwise pile up.
'-----------------------------------------------------------------------
Public Function CloseHandoverWorkbooks() As Long
    Dim folder As String

    folder = HandoverFolder()
    If Len(folder) = 0 Then Exit Function

    CloseHandoverWorkbooks = modUtil.CloseExportWorkbooksUnder(folder)
End Function

'-----------------------------------------------------------------------
' Delete the scratch files -- but only ones old enough that SAP cannot
' still be about to hand them to Excel.
'
' minimumAgeSeconds of 0 means everything, which is what the START of a run
' wants: anything in there is left over from a run that has already ended,
' so no hand-off can be pending. The END of a run passes a grace period
' instead, because the last export may still be in flight; whatever it
' leaves behind is cleared by the next run's opening sweep.
'
' A file that will not delete is one Excel still has open; the next sweep
' gets it. So this reports what it managed and never raises.
'-----------------------------------------------------------------------
Public Function SweepHandover(ByVal minimumAgeSeconds As Double) As Long
    Dim folder As String
    Dim fso As Object
    Dim file As Object
    Dim doomed As Collection
    Dim item As Variant
    Dim age As Double

    folder = HandoverFolder()
    If Len(folder) = 0 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Function

    modUtil.CloseExportWorkbooksUnder folder

    Set doomed = New Collection

    On Error Resume Next
    For Each file In fso.GetFolder(folder).Files
        age = DateDiff("s", file.DateLastModified, Now)
        If age >= minimumAgeSeconds Then doomed.Add file.path
    Next file

    For Each item In doomed
        fso.DeleteFile CStr(item), True
        If Not fso.FileExists(CStr(item)) Then SweepHandover = SweepHandover + 1
    Next item

    ' Fails while anything is still in there, which is fine -- it is out of
    ' the way and the next run reuses it.
    fso.DeleteFolder folder, True
    On Error GoTo 0
End Function

Private Function WaitForFile(ByVal path As String, ByVal maxSeconds As Double) As Boolean
    Dim waited As Double

    Do While waited < maxSeconds
        If modUtil.FileExists(path) Then
            WaitForFile = True
            Exit Function
        End If
        modUtil.SleepSeconds 0.5
        waited = waited + 0.5
    Loop
End Function

' SAP creates the file and then fills it, so 'it exists' is not 'it is
' finished'. Wait until the size stops changing, and answer with it: a copy
' taken mid-write would put a truncated list in the evidence folder and
' nothing downstream would notice.
'
' Still empty counts as settled once it has had a few seconds, so an export
' of an empty list reports that quickly instead of sitting out the whole
' timeout.
Private Function WaitUntilStable(ByVal path As String, ByVal maxSeconds As Double) As Double
    Dim waited As Double
    Dim size As Double, previous As Double

    previous = -1

    Do While waited < maxSeconds
        size = modUtil.FileSizeBytes(path)
        If size = previous And (size > 0 Or waited >= 3) Then
            WaitUntilStable = size
            Exit Function
        End If

        previous = size
        modUtil.SleepSeconds 0.4
        waited = waited + 0.4
    Loop

    WaitUntilStable = modUtil.FileSizeBytes(path)
End Function

' file.txt -> file_2.txt, file_3.txt, ...
Private Function UniqueName(ByVal folder As String, ByVal fileName As String) As String
    Dim stem As String, extension As String
    Dim dotPos As Long
    Dim counter As Long

    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        stem = Left$(fileName, dotPos - 1)
        extension = Mid$(fileName, dotPos)
    Else
        stem = fileName
    End If

    counter = 2
    Do While modUtil.FileExists(modUtil.JoinPath(folder, stem & "_" & counter & extension))
        counter = counter + 1
    Loop

    UniqueName = stem & "_" & counter & extension
End Function

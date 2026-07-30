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
    Dim target As String
    Dim pathId As String, nameId As String

    modUtil.EnsureFolder folder
    target = modUtil.JoinPath(folder, fileName)

    If modUtil.FileExists(target) Then
        fileName = UniqueName(folder, fileName)
        target = modUtil.JoinPath(folder, fileName)
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

    modSapConnect.Element(pathId).Text = folder
    modSapConnect.Element(nameId).Text = fileName

    modSafety.GuardedPress Rebase(modConfig.ElementId("Save.GenerateButton"), windowId)
    modSapConnect.WaitForSap

    If Not WaitForFile(target, 20) Then
        modLog.LogAction sampleIdx, "Export", _
                     "SAP reported no error but " & target & " did not appear.", _
                     "ERROR", target
        Exit Function
    End If

    modLog.LogAction sampleIdx, "Export", _
                 "Wrote " & Format$(modUtil.FileSizeBytes(target) / 1024, "0.0") & " KB", _
                 "OK", target

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
    Dim target As String
    Dim pathId As String, nameId As String, encodingId As String

    If Not modSapConnect.ModalWindowOpen() Then
        modLog.LogAction sampleIdx, "Export", _
                     "No save dialog appeared. The list may be empty, or the export " & _
                     "menu item may be wrong for this release.", "ERROR", vbNullString
        Exit Function
    End If

    modSafety.AssertPopupKnown

    pathId = modConfig.ElementId("Save.Path")
    nameId = modConfig.ElementId("Save.FileName")
    target = modUtil.JoinPath(folder, fileName)

    ' Overwriting silently would destroy evidence from an earlier run.
    If modUtil.FileExists(target) Then
        fileName = UniqueName(folder, fileName)
        target = modUtil.JoinPath(folder, fileName)
    End If

    modSapConnect.Element(pathId).Text = folder

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

    modSapConnect.Element(nameId).Text = fileName

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
    If Not WaitForFile(target, 15) Then
        modLog.LogAction sampleIdx, "Export", _
                     "SAP reported no error but " & target & " did not appear.", _
                     "ERROR", target
        Exit Function
    End If

    If modUtil.FileSizeBytes(target) = 0 Then
        modLog.LogAction sampleIdx, "Export", _
                     "Wrote " & target & " but it is empty -- the list had no rows.", _
                     "ERROR", target
        Exit Function
    End If

    modLog.LogAction sampleIdx, "Export", _
                 "Wrote " & Format$(modUtil.FileSizeBytes(target) / 1024, "0.0") & " KB", _
                 "OK", target

    CompleteSaveDialog = target
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

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
    Dim grid As Object
    Dim target As String
    Dim toolbarButton As String, menuItem As String

    toolbarButton = modConfig.ElementIdOrBlank("Export.AlvToolbarButton")
    menuItem = modConfig.ElementIdOrBlank("Export.AlvMenuItem")

    If Len(toolbarButton) = 0 Or Len(menuItem) = 0 Then
        modLog.LogAction sampleIdx, "Export statement list", _
                     "Skipped -- Export.AlvToolbarButton / Export.AlvMenuItem are blank " & _
                     "on the Screen Map. This export is the FEBAN period view, which is " & _
                     "useful context but not part of the evidence chain, so the run " & _
                     "carries on without it.", "SKIPPED", vbNullString
        Exit Function
    End If

    target = modUtil.JoinPath(folder, fileName)
    If modSafety.BlockedByDryRun("Would export the statement list to " & target) Then Exit Function

    modUtil.EnsureFolder folder
    Set grid = modSapConnect.Element(modConfig.ElementId("FEBAN.ResultGrid"))

    On Error Resume Next
    grid.pressToolbarContextButton toolbarButton
    grid.selectContextMenuItem menuItem
    On Error GoTo 0
    modSapConnect.WaitForSap

    If Not modSapConnect.ModalWindowOpen() Then
        modLog.LogAction sampleIdx, "Export statement list", _
                     "The export menu produced no dialog. Check Export.AlvToolbarButton " & _
                     "and Export.AlvMenuItem against a recording of that right-click.", _
                     "SKIPPED", vbNullString
        Exit Function
    End If

    ConfirmFormatPopupIfPresent
    ExportAlvGrid = CompleteSaveDialog(sampleIdx, folder, fileName)
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
    Dim target As String
    Dim menuId As String

    target = modUtil.JoinPath(folder, fileName)

    If modSafety.BlockedByDryRun("Would export the list to " & target) Then Exit Function

    modUtil.EnsureFolder folder

    menuId = modConfig.ElementIdOrBlank("Export.ListMenu")
    If Len(menuId) > 0 And modSapConnect.Exists(menuId) Then
        modSapConnect.Element(menuId).Select
        modSapConnect.WaitForSap
    Else
        ' Fallback for a list that does offer the OK-code.
        modSafety.GuardedOkCode "%pc"
    End If

    ConfirmFormatPopupIfPresent
    ExportClassicList = CompleteSaveDialog(sampleIdx, folder, fileName)
End Function

' Same save dialog, for callers that opened it themselves -- the invoice
' download in modChain reaches it from the attachment list, not from a list
' export.
Public Function CompleteSaveDialogPublic(ByVal sampleIdx As Long, ByVal folder As String, _
                                         ByVal fileName As String) As String
    modUtil.EnsureFolder folder
    CompleteSaveDialogPublic = CompleteSaveDialog(sampleIdx, folder, fileName)
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

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
' ALV grid export via the toolbar's export button.
' Returns the full path written, or "" if nothing was written.
'-----------------------------------------------------------------------
Public Function ExportAlvGrid(ByVal sampleIdx As Long, ByVal folder As String, _
                              ByVal fileName As String) As String
    Dim grid As Object
    Dim target As String

    target = modUtil.JoinPath(folder, fileName)

    If modSafety.BlockedByDryRun("Would export ALV grid to " & target) Then Exit Function

    modUtil.EnsureFolder folder
    Set grid = modSapConnect.Element(modConfig.ElementId("FEBAN.ResultGrid"))

    On Error Resume Next
    grid.pressToolbarContextButton modConfig.ElementId("Export.ToolbarButton")
    grid.selectContextMenuItem modConfig.ElementId("Export.MenuItem")
    On Error GoTo 0
    modSapConnect.WaitForSap

    ConfirmFormatPopupIfPresent
    ExportAlvGrid = CompleteSaveDialog(sampleIdx, folder, fileName)
End Function

'-----------------------------------------------------------------------
' Classic (non-ALV) list export. %pc is a display-side function.
'-----------------------------------------------------------------------
Public Function ExportClassicList(ByVal sampleIdx As Long, ByVal folder As String, _
                                  ByVal fileName As String) As String
    Dim target As String

    target = modUtil.JoinPath(folder, fileName)

    If modSafety.BlockedByDryRun("Would export classic list to " & target) Then Exit Function

    modUtil.EnsureFolder folder

    modSafety.GuardedOkCode "%pc"
    modSapConnect.WaitForSap

    ConfirmFormatPopupIfPresent
    ExportClassicList = CompleteSaveDialog(sampleIdx, folder, fileName)
End Function

'-----------------------------------------------------------------------
' The 'in which format?' popup. Not every release shows it, and not every
' list offers the same options, so this is best-effort and never fatal.
'-----------------------------------------------------------------------
Private Sub ConfirmFormatPopupIfPresent()
    Dim radioId As String, okId As String

    If Not modSapConnect.ModalWindowOpen() Then Exit Sub

    modSafety.AssertPopupKnown

    radioId = modConfig.ElementIdOrBlank("Export.FormatRadio")
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

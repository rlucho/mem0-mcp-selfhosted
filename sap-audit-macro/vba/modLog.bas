Attribute VB_Name = "modLog"
'=======================================================================
' modLog -- the audit trail.
'
' Every action gets a row, in DRY RUN as well as EXTRACT, stamped with
' the SAP system, client and user it was taken from. An extract without
' this is just a folder of spreadsheets that nobody can vouch for.
'=======================================================================
Option Explicit

Private Const COL_TIMESTAMP As Long = 1
Private Const COL_SAMPLE As Long = 2
Private Const COL_SID As Long = 3
Private Const COL_CLIENT As Long = 4
Private Const COL_USER As Long = 5
Private Const COL_TCODE As Long = 6
Private Const COL_ACTION As Long = 7
Private Const COL_DETAIL As Long = 8
Private Const COL_RESULT As Long = 9
Private Const COL_FILE As Long = 10

Private Const LOG_FIRST_ROW As Long = 5

Public Sub LogAction(ByVal sampleIdx As Long, ByVal action As String, ByVal detail As String, _
                 ByVal result As String, ByVal filePath As String)
    Dim sheet As Worksheet
    Dim row As Long

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_LOG)
    row = NextRow(sheet)

    sheet.Cells(row, COL_TIMESTAMP).Value = Now
    sheet.Cells(row, COL_TIMESTAMP).NumberFormat = "dd/mm/yyyy hh:mm:ss"

    If sampleIdx > 0 Then sheet.Cells(row, COL_SAMPLE).Value = sampleIdx

    sheet.Cells(row, COL_SID).Value = modSapConnect.gSystemId
    sheet.Cells(row, COL_CLIENT).Value = modSapConnect.gClient
    sheet.Cells(row, COL_USER).Value = modSapConnect.gSapUser
    sheet.Cells(row, COL_TCODE).Value = modSapConnect.CurrentTransaction()
    sheet.Cells(row, COL_ACTION).Value = action
    sheet.Cells(row, COL_DETAIL).Value = Left$(detail, 900)
    sheet.Cells(row, COL_RESULT).Value = result
    sheet.Cells(row, COL_FILE).Value = filePath

    Select Case UCase$(result)
        Case "ERROR", "BLOCKED"
            sheet.Cells(row, COL_RESULT).Font.Color = RGB(192, 0, 0)
            sheet.Cells(row, COL_RESULT).Font.Bold = True
        Case "SKIPPED"
            sheet.Cells(row, COL_RESULT).Font.Color = RGB(128, 96, 0)
        Case "OK"
            sheet.Cells(row, COL_RESULT).Font.Color = RGB(0, 112, 48)
    End Select

    ' Flush, so a crash mid-run still leaves the trail on disk.
    Application.StatusBar = "Sample " & sampleIdx & ": " & action
    DoEvents
End Sub

Public Sub WriteHeaderBlock()
    Dim mode As String
    mode = IIf(modConfig.IsDryRun(), "DRY RUN -- nothing will be exported", "EXTRACT")

    LogAction 0, "RUN STARTED", _
          "Mode: " & mode & _
          " | Operator: " & modConfig.Setting("Operator name") & _
          " | Company code: " & modConfig.Setting("Company code") & _
          " | SAP release: " & modSapConnect.gSapRelease & _
          " | Workbook: " & ThisWorkbook.Name, _
          "OK", vbNullString
End Sub

Public Sub WriteFooterBlock(ByVal processed As Long, ByVal errored As Long, ByVal files As Long)
    LogAction 0, "RUN FINISHED", _
          "Samples processed: " & processed & _
          " | Errors: " & errored & _
          " | Files written: " & files & _
          " | Write attempts blocked by the guard: " & modSafety.gWriteAttemptBlocked, _
          IIf(errored = 0, "OK", "ERROR"), vbNullString
End Sub

Private Function NextRow(ByVal sheet As Worksheet) As Long
    Dim lastUsed As Long
    lastUsed = sheet.Cells(sheet.Rows.Count, COL_TIMESTAMP).End(xlUp).Row
    If lastUsed < LOG_FIRST_ROW Then
        NextRow = LOG_FIRST_ROW
    Else
        NextRow = lastUsed + 1
    End If
End Function

' Wipe the trail. Deliberately not called by the run -- a fresh extract
' appends, so successive attempts stay visible. Run it by hand only when
' you mean to discard the history.
Public Sub ClearLog()
    Dim sheet As Worksheet
    Dim lastUsed As Long

    If MsgBox("Delete the entire audit trail on the '" & modConfig.SHEET_LOG & _
              "' sheet?" & vbCrLf & vbCrLf & "This cannot be undone.", _
              vbExclamation + vbYesNo + vbDefaultButton2, "Clear log") <> vbYes Then
        Exit Sub
    End If

    Set sheet = ThisWorkbook.Worksheets(modConfig.SHEET_LOG)
    lastUsed = sheet.Cells(sheet.Rows.Count, COL_TIMESTAMP).End(xlUp).Row
    If lastUsed >= LOG_FIRST_ROW Then
        sheet.Range(sheet.Rows(LOG_FIRST_ROW), sheet.Rows(lastUsed)).Delete
    End If
End Sub

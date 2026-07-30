Attribute VB_Name = "modFeban"
'=======================================================================
' modFeban -- runs the statement search for one month and finds the row
'             matching one sample.
'
' The audit request is per payment line, but FEBAN is per period, so the
' run is organised the other way round: open the month once, then walk
' the samples that fall in it. Ten FEBAN executions instead of 56.
'
' Every element ID comes from the 'Screen Map' sheet. If a step below
' fails with 'element not found', the recorded ID is wrong for this
' screen -- re-record and re-paste rather than editing this module.
'=======================================================================
Option Explicit

' Result of locating a sample in the statement list.
Public Type FebanMatch
    Found As Boolean
    GridRow As Long           ' 0-based, as the ALV addresses its rows
    StatementAmount As Double
    StatementDate As String
    PostingStatus As String
    DocumentNumber As String
    Reference As String
    Ambiguous As Boolean      ' more than one row matched date + amount
    CandidateCount As Long
End Type

'-----------------------------------------------------------------------
' Open FEBAN for one month and execute the selection.
'-----------------------------------------------------------------------
Public Sub OpenMonth(ByVal dateFrom As Date, ByVal dateTo As Date)
    Dim tcode As String
    Dim companyCode As String
    Dim houseBank As String, accountId As String

    tcode = modConfig.Setting("Transaction for statement search")
    companyCode = modConfig.Setting("Company code")
    houseBank = modConfig.Setting("House bank")
    accountId = modConfig.Setting("Account ID")

    modSafety.StartTransaction tcode

    SetField "FEBAN.CompanyCode", companyCode
    SetOptionalField "FEBAN.HouseBank", houseBank
    SetOptionalField "FEBAN.AccountId", accountId
    SetField "FEBAN.StatementDateFrom", modUtil.SapDate(dateFrom)
    SetOptionalField "FEBAN.StatementDateTo", modUtil.SapDate(dateTo)

    modLog.LogAction 0, "FEBAN selection", _
                 "Company code " & companyCode & _
                 ", statement date " & modUtil.SapDate(dateFrom) & _
                 " to " & modUtil.SapDate(dateTo) & _
                 IIf(Len(houseBank) > 0, ", house bank " & houseBank, "") & _
                 IIf(Len(accountId) > 0, ", account " & accountId, ""), _
                 "OK", vbNullString

    modSafety.GuardedPress modConfig.ElementId("FEBAN.ExecuteButton")
    modSafety.AssertPopupKnown

    ' An empty result set is a finding in its own right, not a crash.
    If modSapConnect.StatusBarType() = "E" Or modSapConnect.StatusBarType() = "A" Then
        Err.Raise vbObjectError + 540, "modFeban.OpenMonth", _
                  "FEBAN returned an error for " & modUtil.SapDate(dateFrom) & " to " & _
                  modUtil.SapDate(dateTo) & ": " & modSapConnect.StatusBarText()
    End If
End Sub

Private Sub SetField(ByVal mapKey As String, ByVal value As String)
    modSapConnect.Element(modConfig.ElementId(mapKey)).Text = value
End Sub

' Skip silently when either the ID or the value was left blank -- these
' are the optional FEBAN filters.
Private Sub SetOptionalField(ByVal mapKey As String, ByVal value As String)
    Dim elementId As String

    If Len(value) = 0 Then Exit Sub

    elementId = modConfig.ElementIdOrBlank(mapKey)
    If Len(elementId) = 0 Then Exit Sub
    If Not modSapConnect.Exists(elementId) Then Exit Sub

    modSapConnect.Element(elementId).Text = value
End Sub

'-----------------------------------------------------------------------
' Find the statement row for one sample, by value date and amount.
'
' Matching on date + amount alone is deliberately conservative: where two
' rows tie, the match is reported Ambiguous rather than resolved by
' picking the first. Two of the samples are for identical round amounts
' (2,450,000.00) on month-end dates, so ties are a real prospect.
'-----------------------------------------------------------------------
Public Function FindSample(ByVal paymentDate As Date, ByVal amount As Double) As FebanMatch
    Dim grid As Object
    Dim result As FebanMatch
    Dim rowCount As Long, row As Long
    Dim tolerance As Double
    Dim colDate As String, colAmount As String
    Dim rowDate As String, rowAmount As Double
    Dim wantedDate As String

    Set grid = modSapConnect.Element(modConfig.ElementId("FEBAN.ResultGrid"))
    tolerance = modConfig.SettingNumber("Amount match tolerance", 0.01)
    colDate = modConfig.ElementId("FEBAN.Col.ValueDate")
    colAmount = modConfig.ElementId("FEBAN.Col.Amount")
    wantedDate = modUtil.SapDate(paymentDate)

    rowCount = GridRowCount(grid)
    If rowCount = 0 Then
        result.Found = False
        FindSample = result
        Exit Function
    End If

    For row = 0 To rowCount - 1
        rowDate = Trim$(GridCell(grid, row, colDate))
        rowAmount = modUtil.ParseSapAmount(GridCell(grid, row, colAmount))

        If rowDate = wantedDate Then
            If modUtil.AmountsMatch(rowAmount, amount, tolerance) Then
                result.CandidateCount = result.CandidateCount + 1

                If Not result.Found Then
                    result.Found = True
                    result.GridRow = row
                    result.StatementAmount = rowAmount
                    result.StatementDate = rowDate
                    result.PostingStatus = GridCellIfMapped(grid, row, "FEBAN.Col.Status")
                    result.DocumentNumber = GridCellIfMapped(grid, row, "FEBAN.Col.DocNumber")
                    result.Reference = GridCellIfMapped(grid, row, "FEBAN.Col.Reference")
                Else
                    result.Ambiguous = True
                End If
            End If
        End If
    Next row

    FindSample = result
End Function

'-----------------------------------------------------------------------
' Get back to the statement list after a drill-down.
'
' One F3 is not enough: a drill-down can leave the session two or three
' screens deep, and FB03 is entered as a fresh transaction rather than a
' sub-screen. So press Back until the result grid is on screen again, and
' if that fails, re-execute the selection from scratch.
'
' Returns False when the list could not be restored, which tells the
' caller to stop rather than run the next sample against whatever screen
' happens to be showing.
'-----------------------------------------------------------------------
Public Function ReturnToStatementList(ByVal dateFrom As Date, ByVal dateTo As Date) As Boolean
    Dim gridId As String
    Dim attempt As Long

    gridId = modConfig.ElementId("FEBAN.ResultGrid")

    ' Cheap route first: walk back out of wherever the drill-down landed.
    For attempt = 1 To 4
        If modSapConnect.Exists(gridId) Then
            ReturnToStatementList = True
            Exit Function
        End If

        ' A modal window has to go before Back will do anything.
        If modSapConnect.ModalWindowOpen() Then
            On Error Resume Next
            modSapConnect.Element("wnd[1]").Close
            On Error GoTo 0
            modSapConnect.WaitForSap
        Else
            On Error Resume Next
            modSafety.GuardedSendVKey "wnd[0]", 3
            On Error GoTo 0
        End If
    Next attempt

    If modSapConnect.Exists(gridId) Then
        ReturnToStatementList = True
        Exit Function
    End If

    ' Expensive route: re-run the month's selection. Costs one round trip
    ' per drill-down in the worst case, which is still cheaper than a run
    ' that silently reads the wrong screen.
    On Error GoTo Failed
    OpenMonth dateFrom, dateTo
    ReturnToStatementList = modSapConnect.Exists(gridId)
    Exit Function

Failed:
    modLog.LogAction 0, "Navigation", _
                 "Could not get back to the statement list for " & _
                 Format$(dateFrom, "mmm yy") & ": " & Err.Description, _
                 "ERROR", vbNullString
End Function

' Select the matched row so a drill-down or an export applies to it.
Public Sub SelectRow(ByVal gridRow As Long)
    Dim grid As Object

    Set grid = modSapConnect.Element(modConfig.ElementId("FEBAN.ResultGrid"))

    On Error Resume Next
    grid.selectedRows = CStr(gridRow)
    grid.currentCellRow = gridRow
    On Error GoTo 0

    modSapConnect.WaitForSap
End Sub

'-----------------------------------------------------------------------
' ALV grid access, tolerant of the two shapes these controls come in.
'-----------------------------------------------------------------------
Private Function GridRowCount(ByVal grid As Object) As Long
    On Error Resume Next
    GridRowCount = grid.RowCount
    On Error GoTo 0
End Function

Private Function GridCell(ByVal grid As Object, ByVal row As Long, _
                          ByVal columnName As String) As String
    On Error Resume Next
    GridCell = grid.GetCellValue(row, columnName)
    On Error GoTo 0
End Function

' Returns "" when the column was not mapped on the Screen Map sheet,
' so the optional columns simply come back empty.
Private Function GridCellIfMapped(ByVal grid As Object, ByVal row As Long, _
                                  ByVal mapKey As String) As String
    Dim columnName As String

    columnName = modConfig.ElementIdOrBlank(mapKey)
    If Len(columnName) = 0 Then Exit Function

    GridCellIfMapped = Trim$(GridCell(grid, row, columnName))
End Function

' Column names as this release actually reports them. Run it once against
' a live FEBAN result and paste what it prints into the Screen Map sheet
' -- guessing at FEBAN.Col.* is the commonest reason a run finds nothing.
Public Sub DumpGridColumns()
    Dim grid As Object
    Dim columns As Object
    Dim i As Long
    Dim report As String, title As String

    modConfig.LoadScreenMap
    modSapConnect.SapAttach

    Set grid = modSapConnect.Element(modConfig.ElementId("FEBAN.ResultGrid"))

    On Error Resume Next
    Set columns = grid.ColumnOrder
    On Error GoTo 0

    If columns Is Nothing Then
        MsgBox "This control did not report a column order. It may not be an ALV grid " & _
               "-- check the FEBAN.ResultGrid ID.", vbExclamation
        Exit Sub
    End If

    report = "FEBAN result grid columns (" & columns.Length & ")" & vbCrLf & vbCrLf
    For i = 0 To columns.Length - 1
        ' GetColumnTitles is not on every release, so the technical name is
        ' the part that matters and the readable title is a bonus.
        title = vbNullString
        On Error Resume Next
        title = grid.GetColumnTitles(columns(i))(0)
        On Error GoTo 0

        report = report & columns(i) & "   " & title & vbCrLf
    Next i

    Debug.Print report
    MsgBox report & vbCrLf & "Also written to the Immediate window (Ctrl+G).", _
           vbInformation, "FEBAN grid columns"
End Sub

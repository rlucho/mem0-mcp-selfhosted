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
    Dim selectionWindow As String
    Dim waited As Double

    tcode = modConfig.Setting("Transaction for statement search")
    companyCode = modConfig.Setting("Company code")
    houseBank = modConfig.Setting("House bank")
    accountId = modConfig.Setting("Account ID")
    selectionWindow = modConfig.ElementId("FEBAN.SelectionWindow")

    modSafety.StartTransaction tcode

    ' On this system FEBAN opens its selection as a modal popup, so the fields
    ' do not exist until that window is up. Wait for it rather than assuming.
    Do While Not modSapConnect.Exists(selectionWindow)
        modUtil.SleepSeconds 0.3
        waited = waited + 0.3
        If waited > modConfig.SettingNumber("Max seconds to wait per screen", 60) Then
            Err.Raise vbObjectError + 541, "modFeban.OpenMonth", _
                      "The FEBAN selection window (" & selectionWindow & ") never " & _
                      "appeared after starting " & tcode & ". Current transaction is " & _
                      modSapConnect.CurrentTransaction() & "."
        End If
    Loop

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

    ' The recording pressed a result-screen toolbar button straight after
    ' Execute. It is optional because nobody has confirmed what it does.
    PressPostExecuteIfConfigured

    If modSapConnect.StatusBarType() = "E" Or modSapConnect.StatusBarType() = "A" Then
        Err.Raise vbObjectError + 540, "modFeban.OpenMonth", _
                  "FEBAN returned an error for " & modUtil.SapDate(dateFrom) & " to " & _
                  modUtil.SapDate(dateTo) & ": " & modSapConnect.StatusBarText()
    End If

    If Not modSapConnect.Exists(modConfig.ElementId("FEBAN.ResultGrid")) Then
        Err.Raise vbObjectError + 542, "modFeban.OpenMonth", _
                  "FEBAN ran but the result grid is not on screen. Either the period " & _
                  "holds no statement items, or FEBAN.ResultGrid is wrong for this " & _
                  "screen. Status bar said: " & modSapConnect.StatusBarText()
    End If

    ' Read the whole grid once, here, rather than per sample.
    If LoadGrid(0) = 0 Then
        Err.Raise vbObjectError + 543, "modFeban.OpenMonth", _
                  "The FEBAN result grid reported no rows for " & _
                  modUtil.SapDate(dateFrom) & " to " & modUtil.SapDate(dateTo) & "."
    End If
End Sub

Private Sub PressPostExecuteIfConfigured()
    Dim buttonId As String

    buttonId = modConfig.ElementIdOrBlank("FEBAN.PostExecuteButton")
    If Len(buttonId) = 0 Then Exit Sub
    If Not modSapConnect.Exists(buttonId) Then Exit Sub

    modSafety.GuardedPress buttonId
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
' Cached copy of the whole result grid for the month now open.
'
' WHY THIS EXISTS. A SAP ALV grid only holds the VISIBLE window of rows on
' the client. RowCount reports the true total, but GetCellValue for a row
' outside that window returns nothing -- silently. Walking 0 to RowCount-1
' therefore reads about twenty real rows and then several hundred blanks,
' and every sample past the first screenful comes back NOT FOUND with no
' error to explain it.
'
' The fix is to page through with firstVisibleRow. Doing that once per
' month and caching the result, rather than once per sample, turns roughly
' 25 round trips per sample into 25 per month.
Private mRowDate() As String
Private mRowAmount() As Double
Private mRowStatus() As String
Private mRowDoc() As String
Private mRowRef() As String
Private mRowsLoaded As Long
Private mGridBlanks As Long

Public Sub ClearGridCache()
    Erase mRowDate
    Erase mRowAmount
    Erase mRowStatus
    Erase mRowDoc
    Erase mRowRef
    mRowsLoaded = 0
    mGridBlanks = 0
End Sub

' Page through the result grid and keep every row.
Public Function LoadGrid(ByVal sampleIdx As Long) As Long
    Dim grid As Object
    Dim total As Long, visible As Long
    Dim top As Long, last As Long, row As Long
    Dim colDate As String, colAmount As String
    Dim raw As String

    ClearGridCache

    Set grid = modSapConnect.Element(modConfig.ElementId("FEBAN.ResultGrid"))
    colDate = modConfig.ElementId("FEBAN.Col.ValueDate")
    colAmount = modConfig.ElementId("FEBAN.Col.Amount")

    total = GridRowCount(grid)
    If total = 0 Then Exit Function

    visible = 0
    On Error Resume Next
    visible = grid.VisibleRowCount
    On Error GoTo 0
    If visible <= 0 Then visible = 20

    ReDim mRowDate(0 To total - 1)
    ReDim mRowAmount(0 To total - 1)
    ReDim mRowStatus(0 To total - 1)
    ReDim mRowDoc(0 To total - 1)
    ReDim mRowRef(0 To total - 1)

    top = 0
    Do While top < total
        On Error Resume Next
        grid.firstVisibleRow = top
        On Error GoTo 0
        modSapConnect.WaitForSap

        last = top + visible - 1
        If last > total - 1 Then last = total - 1

        For row = top To last
            raw = GridCell(grid, row, colDate)
            If Len(Trim$(raw)) = 0 Then mGridBlanks = mGridBlanks + 1

            mRowDate(row) = modUtil.GridDateDigits(raw, modUtil.SapDatePatternPublic())
            mRowAmount(row) = modUtil.ParseSapAmount(GridCell(grid, row, colAmount))
            mRowStatus(row) = GridCellIfMapped(grid, row, "FEBAN.Col.Status")
            mRowDoc(row) = GridCellIfMapped(grid, row, "FEBAN.Col.DocNumber")
            mRowRef(row) = GridCellIfMapped(grid, row, "FEBAN.Col.Reference")
        Next row

        top = top + visible
    Loop

    ' Put the grid back to the top, so a later drill-down starts predictably.
    On Error Resume Next
    grid.firstVisibleRow = 0
    On Error GoTo 0

    mRowsLoaded = total

    modLog.LogAction sampleIdx, "Read grid", _
                 "Read " & total & " statement rows in pages of " & visible & _
                 IIf(mGridBlanks > 0, ". " & mGridBlanks & " row(s) came back with no " & _
                     "date -- if that is most of them, FEBAN.Col.ValueDate is wrong.", "."), _
                 IIf(mGridBlanks * 2 > total, "MANUAL", "OK"), vbNullString

    LoadGrid = total
End Function

'-----------------------------------------------------------------------
Public Function FindSample(ByVal paymentDate As Date, ByVal amount As Double) As FebanMatch
    Dim result As FebanMatch
    Dim row As Long
    Dim tolerance As Double
    Dim wantedDate As String

    If mRowsLoaded = 0 Then
        FindSample = result
        Exit Function
    End If

    tolerance = modConfig.SettingNumber("Amount match tolerance", 0.01)

    ' Compared on digits alone. SAP accepts '01092025' typed in but renders
    ' '01.09.2025' in the grid, so a literal comparison would never match --
    ' and that failure is silent, which makes it worth avoiding.
    wantedDate = modUtil.DateDigits(paymentDate)

    For row = 0 To mRowsLoaded - 1
        If mRowDate(row) = wantedDate Then
            If modUtil.AmountsMatch(mRowAmount(row), amount, tolerance) Then
                result.CandidateCount = result.CandidateCount + 1

                If Not result.Found Then
                    result.Found = True
                    result.GridRow = row
                    result.StatementAmount = mRowAmount(row)
                    result.StatementDate = mRowDate(row)
                    result.PostingStatus = mRowStatus(row)
                    result.DocumentNumber = mRowDoc(row)
                    result.Reference = mRowRef(row)
                Else
                    result.Ambiguous = True
                End If
            End If
        End If
    Next row

    FindSample = result
End Function

' What the grid actually held, when a sample did not match. Turns "not
' found" into something diagnosable without another run.
Public Function WhyNoMatch(ByVal paymentDate As Date, ByVal amount As Double) As String
    Dim row As Long
    Dim wantedDate As String
    Dim onDate As Long
    Dim closest As Double, gap As Double, bestGap As Double
    Dim firstDate As String, lastDate As String

    If mRowsLoaded = 0 Then
        WhyNoMatch = "The grid held no readable rows at all."
        Exit Function
    End If

    wantedDate = modUtil.DateDigits(paymentDate)
    bestGap = 1E+30

    For row = 0 To mRowsLoaded - 1
        If Len(mRowDate(row)) = 8 Then
            If Len(firstDate) = 0 Or mRowDate(row) < firstDate Then firstDate = mRowDate(row)
            If mRowDate(row) > lastDate Then lastDate = mRowDate(row)
        End If

        If mRowDate(row) = wantedDate Then
            onDate = onDate + 1
            gap = Abs(Abs(mRowAmount(row)) - Abs(amount))
            If gap < bestGap Then
                bestGap = gap
                closest = mRowAmount(row)
            End If
        End If
    Next row

    If onDate = 0 Then
        WhyNoMatch = mRowsLoaded & " rows read, none dated " & _
                     Format$(paymentDate, "dd/mm/yyyy") & ". Dates present run from " & _
                     Readable(firstDate) & " to " & Readable(lastDate) & _
                     ". If that range looks right, the sample date may be the value date " & _
                     "while the grid column is the statement date."
    Else
        WhyNoMatch = mRowsLoaded & " rows read, " & onDate & " dated " & _
                     Format$(paymentDate, "dd/mm/yyyy") & ", but none for " & _
                     Format$(amount, "#,##0.00") & ". Closest on that date was " & _
                     Format$(closest, "#,##0.00") & "."
    End If
End Function

Private Function Readable(ByVal yyyymmdd As String) As String
    If Len(yyyymmdd) <> 8 Then
        Readable = "(none)"
    Else
        Readable = Mid$(yyyymmdd, 7, 2) & "/" & Mid$(yyyymmdd, 5, 2) & "/" & Left$(yyyymmdd, 4)
    End If
End Function

'-----------------------------------------------------------------------
' Get back to the statement list after a drill-down.
'
' One F3 is not enough: a drill-down can leave the session two or three
' screens deep, and FB03 is entered as a fresh transaction rather than a
' sub-screen. So press Back until the result grid is on screen again, and
' if that fails, re-execute the selection from scratch.
'
' Re-executing also reloads the row cache, which is the expensive case --
' one full paged read of the month. Worth it: the alternative is matching
' later samples against a stale cache.
'-----------------------------------------------------------------------
Public Function ReturnToStatementList(ByVal dateFrom As Date, ByVal dateTo As Date) As Boolean
    Dim gridId As String
    Dim attempt As Long

    gridId = modConfig.ElementId("FEBAN.ResultGrid")

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

' Open the statement item's detail screen, the way the recording does it:
' put the cursor on the amount cell of the row, then double-click it.
Public Sub OpenStatementItem(ByVal gridRow As Long)
    Dim grid As Object
    Dim colAmount As String

    Set grid = modSapConnect.Element(modConfig.ElementId("FEBAN.ResultGrid"))
    colAmount = modConfig.ElementId("FEBAN.Col.Amount")

    grid.setCurrentCell gridRow, colAmount
    modSapConnect.WaitForSap

    grid.doubleClickCurrentCell
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
    modConfig.LoadScreenMap
    modSapConnect.SapAttach
    DumpColumnsOf modConfig.ElementId("FEBAN.ResultGrid"), "FEBAN result grid"
End Sub

' Shared by modFbl1n.DumpGridColumns too.
Public Sub DumpColumnsOf(ByVal gridId As String, ByVal caption As String)
    Dim grid As Object
    Dim columns As Object
    Dim i As Long
    Dim report As String, title As String

    If Not modSapConnect.Exists(gridId) Then
        MsgBox "No control at" & vbCrLf & "  " & gridId & vbCrLf & vbCrLf & _
               "Put the " & caption & " on screen first, then run this again.", _
               vbExclamation, caption
        Exit Sub
    End If

    Set grid = modSapConnect.Element(gridId)

    On Error Resume Next
    Set columns = grid.ColumnOrder
    On Error GoTo 0

    If columns Is Nothing Then
        MsgBox "That control did not report a column order, so it is probably not an " & _
               "ALV grid. Check the ID for " & caption & ".", vbExclamation, caption
        Exit Sub
    End If

    report = caption & " -- " & columns.Length & " columns" & vbCrLf & vbCrLf
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
           vbInformation, caption
End Sub

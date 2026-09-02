Attribute VB_Name = "ResourceLookup"
Option Explicit

'==========================================================
' Resource id lookup   (rev. 2026-09-02)
'
' PURPOSE
'   Fill column P (ProjeQtor resource id) of the "Users Import" sheet
'   automatically, so Update rows stop being skipped.
'
' WHY IT READS AN EXPORT INSTEAD OF COUNTING UP
'   ProjeQtor assigns record ids server-side on INSERT. The workbook
'   cannot know or predict them:
'     * ids are not contiguous -- deleting a record burns its id for good
'       (this instance: activity #636 was deleted, #635 stayed);
'     * anyone else creating a resource in the UI takes the next id
'       without the workbook noticing;
'     * an import that dies partway (max_execution_time) still commits
'       part of its rows, consuming ids.
'   And a WRONG guess is not a harmless miss. Writing an invented id onto
'   a Create row turns an INSERT into an UPDATE of whichever resource
'   already holds that id -- silently overwriting a real person's team,
'   capacity and role. A Create row must carry NO id. That is the whole
'   safety mechanism, so nothing here ever invents one.
'
'   The list is therefore refreshed FROM TRUTH: paste in a ProjeQtor
'   Resource export and the sheet is rebuilt from it. Round trip for a
'   new hire:
'     1. RefreshResourceList          (current export)
'     2. FillResourceIds              -> Create rows stay blank, correctly
'     3. GenerateProjeQtOrCSVs        -> import the resource file
'     4. export Resource again, RefreshResourceList
'        the new hires now have real ids, and every later run updates them
'
' MATCHING KEY: userName (column C), NOT the real name.
'   Of 214 resources, 55 carry a DOUBLE SPACE in the real name
'   ("Francisco  Manzanilla") that HTML collapses on screen, and 11 more
'   are spelled in ways proper-casing destroys ("Silvia De la Fuente
'   Pena", "Cristina Garcia-Rojo Martorell"). userName is unique, ASCII,
'   and is what the person actually logs in with.
'
' The export's column ORDER depends on which fields are on display, so
' columns are found by header label, not by position.
'==========================================================

Private Const LOOKUP_SHEET As String = "Resources"
Private Const IMPORT_SHEET As String = "Users Import"
Private Const DELIM As String = ";"

' Column letters on the LOOKUP_SHEET, written by RefreshResourceList.
'   A id   B userName   C name   D team
Private Const L_ID As Long = 1
Private Const L_USER As Long = 2
Private Const L_NAME As Long = 3
Private Const L_TEAM As Long = 4


'----------------------------------------------------------
' Rebuild the "Resources" sheet from a ProjeQtor Resource export.
'----------------------------------------------------------
Public Sub RefreshResourceList()

    Dim path As String
    path = PickCsv("Select the ProjeQtor Resource export (CSV)")
    If path = "" Then Exit Sub

    Dim lines() As String
    lines = ReadCsvLines(path)
    If UBound(lines) < 1 Then
        MsgBox "That file has no data rows.", vbExclamation
        Exit Sub
    End If

    ' -- locate the columns we need, by label
    Dim head() As String
    head = SplitCsvLine(lines(0))

    Dim cId As Long, cUser As Long, cName As Long, cTeam As Long
    cId = FindCol(head, "id|#|identifiant")
    cUser = FindCol(head, "username|user name|user_name|login|usuario")
    cName = FindCol(head, "name|resource|resource name|nombre|recurso")
    cTeam = FindCol(head, "team|idteam|equipo")

    If cId < 0 Then
        MsgBox "No 'id' column in that export." & vbCrLf & vbCrLf & _
               "Headers found:" & vbCrLf & Join(head, ", "), vbCritical
        Exit Sub
    End If
    If cUser < 0 Then
        MsgBox "No 'userName' column in that export -- add it to the " & _
               "displayed fields in ProjeQtor and export again." & vbCrLf & vbCrLf & _
               "Headers found:" & vbCrLf & Join(head, ", "), vbCritical
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateSheet(LOOKUP_SHEET)
    ws.Cells.Clear
    ws.Cells(1, L_ID).Value = "id"
    ws.Cells(1, L_USER).Value = "userName"
    ws.Cells(1, L_NAME).Value = "name"
    ws.Cells(1, L_TEAM).Value = "team"
    ws.Rows(1).Font.Bold = True

    Dim i As Long, r As Long, f() As String
    Dim maxId As Long, thisId As Long
    Dim dupes As String, seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1                      ' TextCompare

    r = 1
    For i = 1 To UBound(lines)
        If Trim$(lines(i)) <> "" Then
            f = SplitCsvLine(lines(i))
            If UBound(f) >= cId Then
                Dim sId As String, sUser As String
                sId = Trim$(At(f, cId))
                sUser = Trim$(At(f, cUser))
                If sId <> "" Then
                    r = r + 1
                    ' id as TEXT -- it is a key, never arithmetic, and Excel
                    ' would otherwise right-align it and drop leading zeros.
                    ws.Cells(r, L_ID).NumberFormat = "@"
                    ws.Cells(r, L_ID).Value = sId
                    ws.Cells(r, L_USER).Value = sUser
                    ws.Cells(r, L_NAME).Value = At(f, cName)
                    ws.Cells(r, L_TEAM).Value = At(f, cTeam)

                    If IsNumeric(sId) Then
                        thisId = CLng(sId)
                        If thisId > maxId Then maxId = thisId
                    End If

                    If sUser <> "" Then
                        If seen.Exists(sUser) Then
                            dupes = dupes & "  " & sUser & " (ids " & _
                                    seen(sUser) & " and " & sId & ")" & vbCrLf
                        Else
                            seen.Add sUser, sId
                        End If
                    End If
                End If
            End If
        End If
    Next i

    ws.Columns("A:D").AutoFit

    Dim msg As String
    msg = "Resource list refreshed." & vbCrLf & vbCrLf & _
          (r - 1) & " resources" & vbCrLf & _
          "highest id seen: " & maxId & vbCrLf & _
          "source: " & path

    If dupes <> "" Then
        ' Two live resources sharing a userName means FillResourceIds cannot
        ' tell them apart -- it will refuse those rows rather than pick one.
        msg = msg & vbCrLf & vbCrLf & "DUPLICATE userName -- fix in ProjeQtor:" & _
              vbCrLf & dupes
    End If

    MsgBox msg, IIf(dupes = "", vbInformation, vbExclamation)

End Sub


'----------------------------------------------------------
' Fill column P of "Users Import" from the "Resources" sheet.
'
' Only ever WRITES an id it read from the export. A row with no match is
' left blank on purpose: blank means INSERT, which is the right answer
' for someone who does not exist in ProjeQtor yet.
'----------------------------------------------------------
Public Sub FillResourceIds()

    Dim wsL As Worksheet, wsI As Worksheet
    On Error Resume Next
    Set wsL = ThisWorkbook.Worksheets(LOOKUP_SHEET)
    Set wsI = ThisWorkbook.Worksheets(IMPORT_SHEET)
    On Error GoTo 0

    If wsL Is Nothing Then
        MsgBox "No '" & LOOKUP_SHEET & "' sheet yet -- run " & _
               "RefreshResourceList first.", vbExclamation
        Exit Sub
    End If
    If wsI Is Nothing Then
        MsgBox "Sheet '" & IMPORT_SHEET & "' not found.", vbCritical
        Exit Sub
    End If

    ' -- index the lookup by userName, and note duplicates
    Dim byUser As Object, ambiguous As Object
    Set byUser = CreateObject("Scripting.Dictionary")
    Set ambiguous = CreateObject("Scripting.Dictionary")
    byUser.CompareMode = 1
    ambiguous.CompareMode = 1

    Dim lastL As Long, i As Long, k As String
    lastL = wsL.Cells(wsL.Rows.Count, L_ID).End(xlUp).Row
    For i = 2 To lastL
        k = Trim$(CStr(wsL.Cells(i, L_USER).Value))
        If k <> "" Then
            If byUser.Exists(k) Then
                ambiguous(k) = True
            Else
                byUser.Add k, Trim$(CStr(wsL.Cells(i, L_ID).Value))
            End If
        End If
    Next i

    Dim lastI As Long
    lastI = wsI.Cells(wsI.Rows.Count, 2).End(xlUp).Row

    Dim nFilled As Long, nKept As Long, nBlank As Long, nAmbig As Long
    Dim blankList As String, ambigList As String, changedList As String
    Dim action As String, userName As String, existing As String, found As String

    For i = 2 To lastI
        action = Trim$(CStr(wsI.Cells(i, 1).Value))
        userName = Trim$(CStr(wsI.Cells(i, 3).Value))
        existing = Trim$(CStr(wsI.Cells(i, 16).Value))       ' column P
        If action <> "" And userName <> "" Then

            If ambiguous.Exists(userName) Then
                nAmbig = nAmbig + 1
                ambigList = ambigList & "  row " & i & ": " & userName & vbCrLf

            ElseIf byUser.Exists(userName) Then
                found = byUser(userName)
                If existing <> "" And existing <> found Then
                    ' Someone typed an id by hand and it disagrees with the
                    ' export. The export wins, but say so -- a wrong id here
                    ' updates the wrong person.
                    changedList = changedList & "  row " & i & ": " & userName & _
                                  "  " & existing & " -> " & found & vbCrLf
                End If
                wsI.Cells(i, 16).NumberFormat = "@"
                wsI.Cells(i, 16).Value = found
                nFilled = nFilled + 1

            Else
                ' Not in ProjeQtor. Leave P EMPTY so the row INSERTs.
                If existing <> "" Then
                    nKept = nKept + 1
                Else
                    nBlank = nBlank + 1
                    If StrComp(action, "Update", vbTextCompare) = 0 Then
                        blankList = blankList & "  row " & i & ": " & userName & _
                                    "  (marked Update but not in ProjeQtor)" & vbCrLf
                    End If
                End If
            End If
        End If
    Next i

    Dim msg As String
    msg = "Column P filled from '" & LOOKUP_SHEET & "'." & vbCrLf & vbCrLf & _
          nFilled & " matched by userName" & vbCrLf & _
          nBlank & " left blank (not in ProjeQtor -- these INSERT)" & vbCrLf & _
          nKept & " kept a hand-typed id with no export match" & vbCrLf & _
          nAmbig & " skipped, userName not unique"

    If changedList <> "" Then _
        msg = msg & vbCrLf & vbCrLf & "id CORRECTED from the export:" & vbCrLf & changedList
    If ambigList <> "" Then _
        msg = msg & vbCrLf & vbCrLf & "AMBIGUOUS userName -- fill P by hand:" & vbCrLf & ambigList
    If blankList <> "" Then _
        msg = msg & vbCrLf & vbCrLf & "Marked Update but unknown -- GenerateProjeQtOrCSVs " & _
              "will skip these:" & vbCrLf & blankList

    MsgBox msg, vbInformation

End Sub


'==========================================================
' helpers
'==========================================================

Private Function GetOrCreateSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                 After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
    End If
    Set GetOrCreateSheet = ws
End Function


Private Function PickCsv(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .title = title
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV export", "*.csv"
        .Filters.Add "All files", "*.*"
        If .Show = -1 Then PickCsv = .SelectedItems(1) Else PickCsv = ""
    End With
End Function


' Reads the whole file. `Open For Input` decodes as the system ANSI page,
' which on these machines is cp1252 -- the same encoding ProjeQtor exports
' in, so accents survive.
Private Function ReadCsvLines(ByVal path As String) As String()
    Dim fnum As Integer, whole As String
    fnum = FreeFile
    Open path For Input As #fnum
    whole = Input$(LOF(fnum), #fnum)
    Close #fnum
    whole = Replace(whole, vbCrLf, vbLf)
    whole = Replace(whole, vbCr, vbLf)
    ReadCsvLines = Split(whole, vbLf)
End Function


' Splits on ";" but not inside "quotes"; "" is a literal quote.
Private Function SplitCsvLine(ByVal line As String) As String()
    Dim out() As String, n As Long, cur As String
    Dim i As Long, ch As String, inQ As Boolean
    ReDim out(0 To 0)
    For i = 1 To Len(line)
        ch = Mid$(line, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(line, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQ = False
                End If
            Else
                cur = cur & ch
            End If
        ElseIf ch = """" Then
            inQ = True
        ElseIf ch = DELIM Then
            out(n) = cur
            cur = ""
            n = n + 1
            ReDim Preserve out(0 To n)
        Else
            cur = cur & ch
        End If
    Next i
    out(n) = cur
    SplitCsvLine = out
End Function


Private Function At(ByRef f() As String, ByVal idx As Long) As String
    If idx < 0 Then
        At = ""
    ElseIf idx > UBound(f) Then
        At = ""
    Else
        At = Trim$(f(idx))
    End If
End Function


' First header matching any of the "|"-separated aliases, ignoring case,
' spaces and underscores. -1 when absent.
Private Function FindCol(ByRef head() As String, ByVal aliases As String) As Long
    Dim want() As String, i As Long, j As Long, h As String
    want = Split(LCase$(aliases), "|")
    For i = 0 To UBound(head)
        h = LCase$(Replace(Replace(Trim$(head(i)), " ", ""), "_", ""))
        For j = 0 To UBound(want)
            If h = Replace(Replace(want(j), " ", ""), "_", "") Then
                FindCol = i
                Exit Function
            End If
        Next j
    Next i
    FindCol = -1
End Function

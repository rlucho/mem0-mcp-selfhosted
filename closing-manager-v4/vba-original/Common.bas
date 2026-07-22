Public Function SAPsess() As Object

Sheets("config").Range("B1").ClearContents

On Error Resume Next
Set SapGuiAuto = GetObject("SAPGUI")
If Err.Number <> 0 Then
    On Error GoTo 0
    Sheets("config").Range("B1") = "Yes"
    Exit Function
End If
Set SAPapp = SapGuiAuto.GetScriptingEngine
On Error Resume Next
Set Connection = SAPapp.Children(0)
If Err.Number <> 0 Then
    On Error GoTo 0
    Sheets("config").Range("B1") = "Yes"
    Exit Function
End If
On Error GoTo 0
Set SAPsess = Connection.Children(0)

If IsObject(WScript) Then
    WScript.ConnectObject session, "on"
    WScript.ConnectObject SAPapp, "on"
End If

End Function
Function FindLastRow(SearchType As Integer, iColumn As Long, OffsetRows As Long, OffsetColumns As Long, Optional ShName As String)

'find last row for the selected range

'SearchType:
'  1 - selected column
'  2 - entire tab

If SearchType = 1 Then
    If ShName <> "" Then
        FindLastRow = Sheets(ShName).Cells(Rows.Count, iColumn).End(xlUp).offset(OffsetRows, OffsetColumns).Row
    Else
        FindLastRow = Cells(Rows.Count, iColumn).End(xlUp).offset(OffsetRows, OffsetColumns).Row
    End If
ElseIf SearchType = 2 Then
    If ShName <> "" Then
        FindLastRow = Sheets(ShName).Cells.SpecialCells(xlCellTypeLastCell).Row
    Else
        FindLastRow = Cells.SpecialCells(xlCellTypeLastCell).Row
    End If
End If

End Function
Function FindLastColumn(SearchType As Integer, iRow As Long, OffsetRows As Long, OffsetColumns As Long, Optional ShName As String)

'find last column for the selected range

'SearchType:
'  1 - selected row
'  2 - entire tab

If SearchType = 1 Then
    If ShName <> "" Then
        FindLastColumn = Sheets(ShName).Cells(iRow, Columns.Count).End(xlToLeft).offset(OffsetRows, OffsetColumns).Column
    Else
        FindLastColumn = Cells(iRow, Columns.Count).End(xlToLeft).offset(OffsetRows, OffsetColumns).Column
    End If
ElseIf SearchType = 2 Then
    If ShName <> "" Then
        FindLastColumn = Sheets(ShName).Cells.SpecialCells(xlCellTypeLastCell).Column
    Else
        FindLastColumn = Cells.SpecialCells(xlCellTypeLastCell).Column
    End If
End If

End Function
Sub ApplyBorders(BorderType As Integer, ThinType As Integer, rng As Range)

'BorderType:
'  0 - clear borders
'  1 - all borders
'  2 - outside borders

'ThinType:
'  1 - xlThin
'  2 - xlMedium

If ThinType = 1 Then
    WeightType = xlThin
ElseIf ThinType = 2 Then
    WeightType = xlMedium
End If

If BorderType = 0 Then
    With rng
        .Borders(xlEdgeLeft).LineStyle = xlNone
        .Borders(xlEdgeTop).LineStyle = xlNone
        .Borders(xlEdgeBottom).LineStyle = xlNone
        .Borders(xlEdgeRight).LineStyle = xlNone
        .Borders(xlInsideVertical).LineStyle = xlNone
        .Borders(xlInsideHorizontal).LineStyle = xlNone
    End With
ElseIf BorderType = 1 Then
    With rng
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = WeightType
        .Borders(xlEdgeTop).Weight = WeightType
        .Borders(xlEdgeBottom).Weight = WeightType
        .Borders(xlEdgeRight).Weight = WeightType
        .Borders(xlInsideVertical).Weight = WeightType
        .Borders(xlInsideHorizontal).Weight = WeightType
    End With
ElseIf BorderType = 2 Then
    With rng
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Weight = WeightType
        .Borders(xlEdgeTop).Weight = WeightType
        .Borders(xlEdgeBottom).Weight = WeightType
        .Borders(xlEdgeRight).Weight = WeightType
    End With
End If

End Sub

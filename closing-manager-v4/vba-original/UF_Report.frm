Private Sub BT_CANCEL_Click()

Unload Me

End Sub

Private Sub BT_OK_Click()

If CB_CC = "" Then
    MsgBox "You need to select Company Code.", vbCritical
    Exit Sub
End If

Sheets("config").Range("B3").ClearContents
strOption = "Are you sure you want to run closing routine for Company Code " & CB_CC & "?"
With UF_Option
    .Lbl_Option.Caption = strOption
    .StartUpPosition = 0
    .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
    .Show
End With
If Sheets("config").Range("B3") <> "Yes" Then Exit Sub

Sheets("config").Range("B2") = CB_CC

Unload Me

Call RunClosing

End Sub
Private Sub UserForm_Initialize()

LastRow = FindLastRow(1, 10, 0, 0, "config")
If LastRow > 1 Then Sheets("config").Range("J2", Sheets("config").Cells(LastRow, 10)).ClearContents

LastRow = FindLastRow(1, 4, 0, 0, "config")
ArrCC = Sheets("config").Range("D2", Sheets("config").Cells(LastRow, 6))

For i = 1 To UBound(ArrCC, 1)
    If ArrCC(i, 3) = 0 Then
        EmptRow = FindLastRow(1, 10, 1, 0, "config")
        Sheets("config").Cells(EmptRow, 10) = ArrCC(i, 1)
    End If
Next i

Sheets("config").Columns("J:J").RemoveDuplicates Columns:=1, Header:=xlYes

With ActiveWorkbook.Worksheets("config").Sort
    .SortFields.Clear
    .SortFields.Add Key:=Sheets("config").Range("J:J"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    .SetRange Sheets("config").Range("J:J")
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

LastRow = FindLastRow(1, 10, 0, 0, "config")
Sheets("config").Range("J2", Sheets("config").Cells(LastRow, 10)).Name = "CPCList"
CB_CC.RowSource = "CPCList"

End Sub

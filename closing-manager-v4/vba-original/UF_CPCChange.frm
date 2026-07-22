
Option Base 1
Private Sub BT_CANCEL_Click()

Unload Me

End Sub

Private Sub BT_OK_Click()

If CB_CC = "" Then
    MsgBox "You need to select the Company Code and make the changes.", vbCritical
    Exit Sub
End If

'check if any changes were entered
CC = CB_CC.Value
LastRow = FindLastRow(1, 4, 0, 0, "config")
ArrCPC = Sheets("config").Range("C2", Sheets("config").Cells(LastRow, 9))
AnythingChanged = False
For i = 1 To UBound(ArrCPC, 1)
    If CC = ArrCPC(i, 2) Then
        IDn = ArrCPC(i, 1)
        If (ChB_LocationClosed.Value = False And ArrCPC(i, 4) = 1) Or (ChB_LocationClosed.Value = True And ArrCPC(i, 4) = 0) Then
            AnythingChanged = True
            Exit For
        End If
        If (ChB_CPCAllowedGAAP.Value = False And ArrCPC(i, 5) = 1) Or (ChB_CPCAllowedGAAP.Value = True And ArrCPC(i, 5) = 0) Then
            AnythingChanged = True
            Exit For
        End If
        If (ChB_CPCNotAllowed.Value = False And ArrCPC(i, 6) = 1) Or (ChB_CPCNotAllowed.Value = True And ArrCPC(i, 6) = 0) Then
            AnythingChanged = True
            Exit For
        End If
        If TB_Comment <> ArrCPC(i, 7) Then
            AnythingChanged = True
            Exit For
        End If
        Exit For
    End If
Next i
If AnythingChanged = False Then
    MsgBox "Nothing was changed. Click 'Cancel' to close the form.", vbInformation
    Exit Sub
End If

odp = MsgBox("Are you sure you want to make the changes?", vbYesNo + vbQuestion)
If odp = vbNo Then Exit Sub

'implement the changes
CC = CB_CC.Value
LastRow = FindLastRow(1, 4, 0, 0, "config")
ArrCPC = Sheets("config").Range("D2", Sheets("config").Cells(LastRow, 9))
For i = 1 To UBound(ArrCPC, 1)
    If CC = ArrCPC(i, 1) Then
        If ChB_LocationClosed.Value = True Then ArrCPC(i, 3) = 1 Else ArrCPC(i, 3) = 0
        If ChB_CPCAllowedGAAP.Value = True Then ArrCPC(i, 4) = 1 Else ArrCPC(i, 4) = 0
        If ChB_CPCNotAllowed.Value = True Then ArrCPC(i, 5) = 1 Else ArrCPC(i, 5) = 0
        ArrCPC(i, 6) = TB_Comment
        Exit For
    End If
Next i

Sheets("config").Range("D2", Sheets("config").Cells(LastRow, 9)) = ArrCPC

'upload the changes to sharepoint

List = "CCCrossList"

updates = "<Batch> <Method ID='1' Cmd='Update'>" & _
                            "<Field Name='ID'>" & IDn & "</Field>" & _
                            "<Field Name='LocationClosed'>" & ArrCPC(i, 3) & "</Field>" & _
                            "<Field Name='CPCAllowedGAAP'>" & ArrCPC(i, 4) & "</Field>" & _
                            "<Field Name='CPCNotAllowed'>" & ArrCPC(i, 5) & "</Field>" & _
                            "<Field Name='Comment'>" & ArrCPC(i, 6) & "</Field>" & _
                            "</Method></Batch>"

Call spUpdateList(updates, List)

Unload Me

End Sub
Private Sub CB_CC_Change()

CC = CB_CC.Value
LastRow = FindLastRow(1, 4, 0, 0, "config")
ArrCPC = Sheets("config").Range("D2", Sheets("config").Cells(LastRow, 9))

For i = 1 To UBound(ArrCPC, 1)
    If CC = ArrCPC(i, 1) Then
        If ArrCPC(i, 3) = 0 Then ChB_LocationClosed.Value = False Else ChB_LocationClosed.Value = True
        If ArrCPC(i, 4) = 0 Then ChB_CPCAllowedGAAP.Value = False Else ChB_CPCAllowedGAAP.Value = True
        If ArrCPC(i, 5) = 0 Then ChB_CPCNotAllowed.Value = False Else ChB_CPCNotAllowed.Value = True
        TB_Comment = ArrCPC(i, 6)
        Exit For
    End If
Next i

End Sub
Private Sub ChB_CPCAllowedGAAP_Click()

If ChB_CPCAllowedGAAP.Value = True Then
    ChB_CPCNotAllowed.Value = False
End If

End Sub

Private Sub ChB_CPCNotAllowed_Click()

If ChB_CPCNotAllowed.Value = True Then
    ChB_CPCAllowedGAAP.Value = False
End If

End Sub

Private Sub UserForm_Initialize()

LastRow = FindLastRow(1, 3, 0, 0, "config")
Sheets("config").Range("D2", Sheets("config").Cells(LastRow, 4)).Name = "CPCList"
CB_CC.RowSource = "CPCList"

End Sub

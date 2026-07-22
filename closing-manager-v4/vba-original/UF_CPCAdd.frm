Option Base 1
Private Sub BT_CANCEL_Click()

Unload Me

End Sub

Private Sub BT_OK_Click()

If TB_CC = "" Then
    MsgBox "You need to provide Company Code.", vbCritical
    Exit Sub
End If
If Len(TB_CC) <> 4 Or IsNumeric(Right(TB_CC, 2)) = False Or IsNumeric(Left(TB_CC, 1)) = True Or IsNumeric(Mid(TB_CC, 2, 1)) = True Then
    MsgBox "You need to provide correct Company Code.", vbCritical
    Exit Sub
End If

'check SAP connection
Set sess = SAPsess

If Sheets("config").Range("B1") = "Yes" Then
    MsgBox "You are not logged in SAP. Please log in and run macro again.", vbCritical
    Exit Sub
End If

Call CreatePaths

'check if Company Code already on the list
CC = TB_CC
LastRow = FindLastRow(1, 3, 0, 0, "config")
ArrCPC = Sheets("config").Range("C2", Sheets("config").Cells(LastRow, 8))
found = False
For i = 1 To UBound(ArrCPC, 1)
    If CC = ArrCPC(i, 1) Then
        found = True
        Exit For
    End If
Next i

If found = True Then
    MsgBox "Company Code already exists on the list.", vbCritical
    Exit Sub
End If

odp = MsgBox("Are you sure you want to add the Company Code with selected settings?", vbYesNo + vbQuestion)
If odp = vbNo Then Exit Sub

'implement the changes
EmptRow = FindLastRow(1, 3, 1, 0, "config")
Sheets("config").Cells(EmptRow, 3) = TB_CC
If ChB_LocationClosed.Value = True Then Sheets("config").Cells(EmptRow, 5) = True Else Sheets("config").Cells(EmptRow, 5) = False
If ChB_CPCAllowedGAAP.Value = True Then Sheets("config").Cells(EmptRow, 6) = True Else Sheets("config").Cells(EmptRow, 6) = False
If ChB_CPCNotAllowed.Value = True Then Sheets("config").Cells(EmptRow, 7) = True Else Sheets("config").Cells(EmptRow, 7) = False
Sheets("config").Cells(EmptRow, 8) = TB_Comment

'check the posting block on 5100001 account


'upload the changes to sharepoint

Unload Me

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

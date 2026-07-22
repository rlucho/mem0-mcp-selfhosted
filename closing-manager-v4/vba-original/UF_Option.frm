Private Sub BT_NO_Click()
Sheets("config").Range("B3") = "No"
Unload Me
End Sub
Private Sub BT_YES_Click()
Sheets("config").Range("B3") = "Yes"
Unload Me
End Sub

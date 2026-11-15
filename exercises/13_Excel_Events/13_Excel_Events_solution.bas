Attribute VB_Name = "Exercise"
Private Sub Workbook_Open()
    MsgBox "Welcome to VBA Training!"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        MsgBox "Cell A1 changed to: " & Target.Value
    End If
End Sub

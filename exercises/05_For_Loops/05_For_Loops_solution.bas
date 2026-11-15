Attribute VB_Name = "Exercise"
Sub ForLoops()
    Dim i As Integer
    Dim output As String
    
    output = ""
    
    For i = 1 To 10
        output = output & i & vbCrLf
        If i = 7 Then
            Exit For
        End If
    Next i
    
    MsgBox output
End Sub

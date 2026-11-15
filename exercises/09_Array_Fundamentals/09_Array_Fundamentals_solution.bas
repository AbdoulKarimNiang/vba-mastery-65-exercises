Attribute VB_Name = "Exercise"
Sub ArrayFundamentals()
    Dim myArray(1 To 5) As Integer
    Dim i As Integer
    Dim output As String
    
    For i = 1 To 5
        myArray(i) = i * 10
    Next i
    
    output = ""
    For i = 1 To 5
        output = output & myArray(i) & vbCrLf
    Next i
    
    MsgBox "Array contents:" & vbCrLf & output
End Sub

Attribute VB_Name = "Exercise"
Sub DoWhileLoops()
    Dim count As Integer
    count = 0
    
    Do While count < 10
        Debug.Print count
        count = count + 1
        If count = 5 Then
            Exit Do
        End If
    Loop
    
    MsgBox "Loop completed at count: " & count
End Sub

Attribute VB_Name = "Exercise"
Sub ConditionalLogic()
    Dim score As Integer
    Dim grade As String
    
    score = 85
    
    If score >= 90 Then
        grade = "A"
    ElseIf score >= 80 Then
        grade = "B"
    ElseIf score >= 70 Then
        grade = "C"
    Else
        grade = "F"
    End If
    
    MsgBox "Score: " & score & " Grade: " & grade
End Sub

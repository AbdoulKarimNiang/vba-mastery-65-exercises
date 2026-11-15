# VBA Exercise 7: If ElseIf Else — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `07_If_ElseIf_Else_solution.xlsm` and `07_If_ElseIf_Else_solution.bas`.

## Solution code

```vba
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
```

## Explanation

**Select Case Alternative:**
```vba
Select Case score
    Case 90 To 100
        grade = "A"
    Case 80 To 89
        grade = "B"
    Case 70 To 79
        grade = "C"
    Case Else
        grade = "F"
End Select
```

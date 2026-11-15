# VBA Exercise 5: For Loops — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `05_For_Loops_solution.xlsm` and `05_For_Loops_solution.bas`.

## Solution code

```vba
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
```

## Explanation

**Loop Variations:**
```vba
' Count down
For i = 10 To 1 Step -1

' Count by 2s
For i = 1 To 10 Step 2

' Each cell in range
For Each cell In Range("A1:A10")

' Each worksheet
For Each sheet In ThisWorkbook.Sheets
```

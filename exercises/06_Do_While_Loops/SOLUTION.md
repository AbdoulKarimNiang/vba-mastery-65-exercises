# VBA Exercise 6: Do While Loops — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `06_Do_While_Loops_solution.xlsm` and `06_Do_While_Loops_solution.bas`.

## Solution code

```vba
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
```

## Explanation

**Do Loop Variations:**
```vba
' Do While - checks at start
Do While condition
    ' code
Loop

' Do Until - opposite of While
Do Until condition
    ' code
Loop

' Check at end
Do
    ' code
Loop While condition
```

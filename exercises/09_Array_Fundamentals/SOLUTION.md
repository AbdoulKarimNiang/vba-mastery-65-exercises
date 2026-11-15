# VBA Exercise 9: Array Fundamentals — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `09_Array_Fundamentals_solution.xlsm` and `09_Array_Fundamentals_solution.bas`.

## Solution code

```vba
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
```

## Explanation

**Array Operations:**
- `Dim arr(1 To 10)` — Fixed array
- `Dim arr()` — Dynamic array
- `ReDim arr(1 To 20)` — Resize
- `ReDim Preserve arr()` — Keep existing data
- `LBound(arr)` — Lower bound index
- `UBound(arr)` — Upper bound index

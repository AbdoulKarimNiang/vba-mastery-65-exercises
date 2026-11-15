# VBA Exercise 8: Procedures Functions — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `08_Procedures_Functions_solution.xlsm` and `08_Procedures_Functions_solution.bas`.

## Solution code

```vba
Sub CreateFunctions()
    Dim radius As Double
    Dim celsius As Double
    
    radius = 5
    celsius = 25
    
    MsgBox "Circle area: " & CircleArea(radius) & vbCrLf & _
           "25C in F: " & CelsiusToFahrenheit(celsius)
End Sub

Function CircleArea(r As Double) As Double
    CircleArea = 3.14159 * r * r
End Function

Function CelsiusToFahrenheit(c As Double) As Double
    CelsiusToFahrenheit = (c * 9/5) + 32
End Function
```

## Explanation

**Key Differences:**
- `Sub` — Does something, returns nothing
- `Function` — Does something, returns a value
- Both can have parameters
- Return value = assign to function name

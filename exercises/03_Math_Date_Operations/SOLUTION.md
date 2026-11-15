# VBA Exercise 3: Math Date Operations — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `03_Math_Date_Operations_solution.xlsm` and `03_Math_Date_Operations_solution.bas`.

## Solution code

```vba
Sub MathDateOperations()
    Dim price As Double
    Dim tax As Double
    Dim totalPrice As Double
    Dim today As Date
    Dim endOfYear As Date
    Dim daysRemaining As Integer
    
    price = 100
    tax = 0.15
    totalPrice = price + (price * tax)
    
    today = Date
    endOfYear = DateSerial(Year(today), 12, 31)
    daysRemaining = endOfYear - today
    
    MsgBox "Price: $" & Format(price, "0.00") & vbCrLf & _
           "Total with tax: $" & Format(totalPrice, "0.00") & vbCrLf & _
           "Days until year end: " & daysRemaining
End Sub
```

## Explanation

**Key Functions:**
- `Date` — Today's date
- `DateSerial(year, month, day)` — Create date
- `DateAdd("m", 3, date)` — Add to date
- `Year()`, `Month()`, `Day()` — Extract parts
- `Format(value, "0.00")` — Format as decimal
- Math: `+`, `-`, `*`, `/`, `^` (power), `Mod` (remainder)

**Date Arithmetic:**
- Subtracting dates gives days between
- Adding/subtracting numbers from dates works

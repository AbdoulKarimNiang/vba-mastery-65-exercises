# VBA Exercise 4: Working With Ranges — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made code: `04_Working_With_Ranges_solution.bas` (import it with **File > Import File...** in the VBA editor).

## Solution code

```vba
Sub WorkingWithRanges()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    Dim sumValue As Double
    Dim countValue As Long
    Dim avgValue As Double

    Set ws = ThisWorkbook.Worksheets("Data")
    Set dataRange = ws.Range("B2:F2")

    For Each cell In dataRange.Cells
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
            sumValue = sumValue + cell.Value
            countValue = countValue + 1
        End If
    Next cell

    If countValue > 0 Then avgValue = sumValue / countValue

    ws.Range("H1").Value = sumValue
    ws.Range("H2").Value = avgValue

    MsgBox "Sum: " & sumValue & vbCrLf & "Average: " & avgValue
End Sub
```

Expected result: **H1 = 150**, **H2 = 30**.

## How it works

- `ws.Range("B2:F2")` gets the five values in row 2 as one `Range` object.
- `For Each cell In dataRange.Cells` goes through each cell in that range.
- `IsEmpty` and the built-in `IsNumeric` skip blanks and text, so a stray label doesn't cause a
  *Type mismatch* error.
- The average divides by `countValue` (the number of cells actually added), not a hard-coded 5,
  so skipped cells don't skew it.
- Results go to H1:H2, away from the data in A1:F2.

### Shorter alternative

Excel's worksheet functions do the same work in two lines, and ignore blanks and text on their own:

```vba
ws.Range("H1").Value = Application.WorksheetFunction.Sum(dataRange)
ws.Range("H2").Value = Application.WorksheetFunction.Average(dataRange)
```

## Range access methods

- `Range("A1")`: a single cell
- `Range("B2:F2")`: a block of cells
- `Cells(row, column)`: a cell by row/column number, e.g. `Cells(2, 3)` is C2
- `someRange.Offset(rows, cols)`: a cell relative to another, e.g. `Range("A1").Offset(1, 2)` is C2
- `.Value`: get or set a cell's value
- `.Address`: get a cell's address as text, e.g. `"$B$2"`

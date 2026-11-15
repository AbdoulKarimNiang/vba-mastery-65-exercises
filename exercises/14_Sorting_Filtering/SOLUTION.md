# VBA Exercise 14: Sorting Filtering — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `14_Sorting_Filtering_solution.xlsm` and `14_Sorting_Filtering_solution.bas`.

## Solution code

```vba
Sub SortAndFilter()
    Dim ws As Worksheet
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Sheets("Data")
    Set dataRange = ws.Range("A1:B11")
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B1"), Order:=xlDescending
        .SetRange dataRange
        .Header = xlYes
        .Apply
    End With
    
    dataRange.AutoFilter
    MsgBox "Sorted and filtered"
End Sub
```

## Explanation

**Sort Options:**
- `xlAscending` — A to Z
- `xlDescending` — Z to A
- `.Header = xlYes` — Has headers

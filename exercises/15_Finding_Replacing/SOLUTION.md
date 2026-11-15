# VBA Exercise 15: Finding Replacing — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `15_Finding_Replacing_solution.xlsm` and `15_Finding_Replacing_solution.bas`.

## Solution code

```vba
Sub FindReplace()
    Dim ws As Worksheet
    Dim foundCell As Range
    Dim replaceCount As Integer
    
    Set ws = ThisWorkbook.Sheets("Data")
    replaceCount = 0
    
    Set foundCell = ws.Columns(1).Find(What:="old", LookAt:=xlWhole)
    
    Do While Not foundCell Is Nothing
        foundCell.Value = "new"
        replaceCount = replaceCount + 1
        Set foundCell = ws.Columns(1).Find(What:="old", After:=foundCell)
    Loop
    
    MsgBox "Replaced " & replaceCount & " occurrences"
End Sub
```

## Explanation

**Find Method:**
- `.Find(What:=..., LookAt:=...)` — Search
- `LookAt:=xlWhole` — Exact match
- `LookAt:=xlPart` — Partial match
- Loop through all matches with While Not foundCell Is Nothing

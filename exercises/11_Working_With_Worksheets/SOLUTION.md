# VBA Exercise 11: Working With Worksheets — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `11_Working_With_Worksheets_solution.xlsm` and `11_Working_With_Worksheets_solution.bas`.

## Solution code

```vba
Sub ManageWorksheets()
    Dim newSheet As Worksheet
    Dim originalSheet As Worksheet
    
    Set originalSheet = ThisWorkbook.ActiveSheet
    
    Set newSheet = ThisWorkbook.Sheets.Add
    newSheet.Name = "Report"
    newSheet.Range("A1").Value = "Hello"
    
    originalSheet.Activate
    MsgBox "Created new sheet: " & newSheet.Name
End Sub
```

## Explanation

**Worksheet Operations:**
- `Sheets.Add` — Create new sheet
- `.Delete` — Remove sheet
- `.Name` — Rename
- `.Activate` — Switch to sheet
- `.Copy` — Copy sheet
- `Worksheets("name")` — Access by name
- `Worksheets(1)` — Access by index

# VBA Exercise 11: Working With Worksheets

## Objective

Manage multiple sheets and navigate between them

## Files

| File | Purpose |
|------|---------|
| `11_Working_With_Worksheets_start.xlsm` | Workbook to work in |
| `11_Working_With_Worksheets_start.bas` | Starting code to import |
| `SOLUTION.md` | Solution and explanation (open only when you're done) |

## How to load the starting code

1. Open `11_Working_With_Worksheets_start.xlsm` and click **Enable Content** if asked.
2. Press **Alt+F11** to open the VBA editor.
3. **File > Import File...** and choose `11_Working_With_Worksheets_start.bas`.
4. Complete the `TODO`s, run with **F5**, and save with **Ctrl+S**.

> Event procedures (`Workbook_Open`, `Worksheet_Change`) must be pasted into the
> **ThisWorkbook** or sheet code module, not imported as a standard module.

## Starting code

```vba
Sub ManageWorksheets()
    ' TODO: Add a new sheet named "Report"
    ' TODO: Write "Hello" to cell A1
    ' TODO: Switch to original sheet
End Sub
```

## Check your work

When your code runs correctly, compare it with [SOLUTION.md](SOLUTION.md)
or the `11_Working_With_Worksheets_solution.xlsm` / `11_Working_With_Worksheets_solution.bas` files.

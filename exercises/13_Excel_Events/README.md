# VBA Exercise 13: Excel Events

## Objective

Respond to Excel events like Workbook_Open and Worksheet_Change

## Files

| File | Purpose |
|------|---------|
| `13_Excel_Events_start.xlsm` | Workbook to work in |
| `13_Excel_Events_start.bas` | Starting code to import |
| `SOLUTION.md` | Solution and explanation (open only when you're done) |

## How to load the starting code

1. Open `13_Excel_Events_start.xlsm` and click **Enable Content** if asked.
2. Press **Alt+F11** to open the VBA editor.
3. **File > Import File...** and choose `13_Excel_Events_start.bas`.
4. Complete the `TODO`s, run with **F5**, and save with **Ctrl+S**.

> Event procedures (`Workbook_Open`, `Worksheet_Change`) must be pasted into the
> **ThisWorkbook** or sheet code module, not imported as a standard module.

## Starting code

```vba
Private Sub Workbook_Open()
    ' TODO: Display welcome message when workbook opens
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    ' TODO: Show message when cell A1 is changed
End Sub
```

## Check your work

When your code runs correctly, compare it with [SOLUTION.md](SOLUTION.md)
or the `13_Excel_Events_solution.xlsm` / `13_Excel_Events_solution.bas` files.

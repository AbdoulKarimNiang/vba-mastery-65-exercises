# VBA Exercise 7: If ElseIf Else

## Objective

Implement conditional logic with If-ElseIf-Else statements

## Files

| File | Purpose |
|------|---------|
| `07_If_ElseIf_Else_start.xlsm` | Workbook to work in |
| `07_If_ElseIf_Else_start.bas` | Starting code to import |
| `SOLUTION.md` | Solution and explanation (open only when you're done) |

## How to load the starting code

1. Open `07_If_ElseIf_Else_start.xlsm` and click **Enable Content** if asked.
2. Press **Alt+F11** to open the VBA editor.
3. **File > Import File...** and choose `07_If_ElseIf_Else_start.bas`.
4. Complete the `TODO`s, run with **F5**, and save with **Ctrl+S**.

> Event procedures (`Workbook_Open`, `Worksheet_Change`) must be pasted into the
> **ThisWorkbook** or sheet code module, not imported as a standard module.

## Starting code

```vba
Sub ConditionalLogic()
    Dim score As Integer
    Dim grade As String
    
    ' TODO: Input score and assign grade
    ' 90+ = A, 80-89 = B, 70-79 = C, <70 = F
End Sub
```

## Check your work

When your code runs correctly, compare it with [SOLUTION.md](SOLUTION.md)
or the `07_If_ElseIf_Else_solution.xlsm` / `07_If_ElseIf_Else_solution.bas` files.

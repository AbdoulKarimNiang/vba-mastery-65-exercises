# VBA Exercise 1: Variables Basics

## Objective

Declare and use different variable types (Integer, String, Double, Boolean)

## Files

| File | Purpose |
|------|---------|
| `01_Variables_Basics_start.xlsm` | Workbook to work in |
| `01_Variables_Basics_start.bas` | Starting code to import |
| `SOLUTION.md` | Solution and explanation (open only when you're done) |

## How to load the starting code

1. Open `01_Variables_Basics_start.xlsm` and click **Enable Content** if asked.
2. Press **Alt+F11** to open the VBA editor.
3. **File > Import File...** and choose `01_Variables_Basics_start.bas`.
4. Complete the `TODO`s, run with **F5**, and save with **Ctrl+S**.

> Event procedures (`Workbook_Open`, `Worksheet_Change`) must be pasted into the
> **ThisWorkbook** or sheet code module, not imported as a standard module.

## Starting code

```vba
Sub DeclareVariables()
    ' TODO: Declare these variables:
    ' - myAge as Integer with value 25
    ' - myName as String with value "John"
    ' - myHeight as Double with value 5.9
    ' - isActive as Boolean with value True
    
    ' Then display them in a message box
End Sub
```

## Check your work

When your code runs correctly, compare it with [SOLUTION.md](SOLUTION.md)
or the `01_Variables_Basics_solution.xlsm` / `01_Variables_Basics_solution.bas` files.

# VBA Exercise 2: String Operations

## Objective

Concatenate, format, and manipulate strings using VBA functions

## Files

| File | Purpose |
|------|---------|
| `02_String_Operations_start.xlsm` | Workbook to work in |
| `02_String_Operations_start.bas` | Starting code to import |
| `SOLUTION.md` | Solution and explanation (open only when you're done) |

## How to load the starting code

1. Open `02_String_Operations_start.xlsm` and click **Enable Content** if asked.
2. Press **Alt+F11** to open the VBA editor.
3. **File > Import File...** and choose `02_String_Operations_start.bas`.
4. Complete the `TODO`s, run with **F5**, and save with **Ctrl+S**.

> Event procedures (`Workbook_Open`, `Worksheet_Change`) must be pasted into the
> **ThisWorkbook** or sheet code module, not imported as a standard module.

## Starting code

```vba
Sub StringOperations()
    Dim firstName As String
    Dim lastName As String
    ' TODO: Concatenate first and last name
    ' TODO: Use UPPER(), LOWER(), LEN() functions
    ' TODO: Extract first 3 characters using LEFT()
End Sub
```

## Check your work

When your code runs correctly, compare it with [SOLUTION.md](SOLUTION.md)
or the `02_String_Operations_solution.xlsm` / `02_String_Operations_solution.bas` files.

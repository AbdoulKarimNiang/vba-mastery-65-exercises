# VBA Exercise 4: Working With Ranges

## Objective

Read from and write to Excel ranges using VBA

## Files

| File | Purpose |
|------|---------|
| `04_Working_With_Ranges_start.xlsm` | Workbook to work in |
| `04_Working_With_Ranges_start.bas` | Starting code to import |
| `SOLUTION.md` | Solution and explanation (open only when you're done) |

## The data

The `Data` sheet holds its values in a row, not a column:

|   | A      | B  | C  | D  | E  | F  |
|---|--------|----|----|----|----|----|
| 1 | Number | 1  | 2  | 3  | 4  | 5  |
| 2 | Value  | 10 | 20 | 30 | 40 | 50 |

Sum the five values in **B2:F2** and write the results to **H1** (sum) and **H2** (average),
outside the data so nothing gets overwritten.

## How to load the starting code

1. Open `04_Working_With_Ranges_start.xlsm` and click **Enable Content** if asked.
2. Press **Alt+F11** to open the VBA editor.
3. **File > Import File...** and choose `04_Working_With_Ranges_start.bas`.
4. Complete the `TODO`s and run with **F5**.
5. Save with **Ctrl+S**, keeping the **.xlsm** (macro-enabled) format so your code is saved.

## Starting code

```vba
Sub WorkingWithRanges()
    ' Sheet: Data has headers in A1:A2 and values in B2:F2
    ' TODO: Read values from range B2:F2
    ' TODO: Calculate sum and average
    ' TODO: Write results to cells H1 and H2
End Sub
```

## Check your work

Expected result: **H1 = 150** and **H2 = 30**.

Then compare your code with [SOLUTION.md](SOLUTION.md) or `04_Working_With_Ranges_solution.bas`.

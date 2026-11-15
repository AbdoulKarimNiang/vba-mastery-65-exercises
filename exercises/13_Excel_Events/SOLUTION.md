# VBA Exercise 13: Excel Events — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `13_Excel_Events_solution.xlsm` and `13_Excel_Events_solution.bas`.

## Solution code

```vba
Private Sub Workbook_Open()
    MsgBox "Welcome to VBA Training!"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    If Target.Address = "$A$1" Then
        MsgBox "Cell A1 changed to: " & Target.Value
    End If
End Sub
```

## Explanation

**Common Events:**
- `Workbook_Open()` — File opens
- `Worksheet_Change(Target)` — Cell changes
- `Worksheet_Activate()` — Sheet activated
- `Worksheet_Deactivate()` — Sheet left

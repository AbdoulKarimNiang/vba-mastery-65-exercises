# VBA Exercise 10: Splitting Joining Data — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `10_Splitting_Joining_Data_solution.xlsm` and `10_Splitting_Joining_Data_solution.bas`.

## Solution code

```vba
Sub SplitJoinData()
    Dim dataString As String
    Dim parts() As String
    Dim newString As String
    
    dataString = "John,Smith,25,New York"
    parts = Split(dataString, ",")
    
    MsgBox "First name: " & parts(0) & vbCrLf & _
           "City: " & parts(3) & vbCrLf & _
           "Joined: " & Join(parts, ";")
End Sub
```

## Explanation

**Key Functions:**
- `Split(string, delimiter)` — Returns array
- `Join(array, delimiter)` — Returns string
- Array indexing starts at 0 (not 1!)
- UBound() helps find last element

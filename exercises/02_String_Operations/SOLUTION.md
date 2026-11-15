# VBA Exercise 2: String Operations — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `02_String_Operations_solution.xlsm` and `02_String_Operations_solution.bas`.

## Solution code

```vba
Sub StringOperations()
    Dim firstName As String
    Dim lastName As String
    Dim fullName As String
    Dim upperName As String
    Dim lowerName As String
    Dim nameLength As Integer
    Dim firstThree As String
    
    firstName = "John"
    lastName = "Smith"
    
    fullName = firstName & " " & lastName
    upperName = UCase(fullName)
    lowerName = LCase(fullName)
    nameLength = Len(fullName)
    firstThree = Left(fullName, 3)
    
    MsgBox "Full: " & fullName & vbCrLf & _
           "Upper: " & upperName & vbCrLf & _
           "Lower: " & lowerName & vbCrLf & _
           "Length: " & nameLength & vbCrLf & _
           "First 3: " & firstThree
End Sub
```

## Explanation

**Key Functions:**
- `&` — Concatenate strings
- `UCase()` — Convert to UPPERCASE
- `LCase()` — Convert to lowercase
- `Len()` — String length
- `Left(string, n)` — First n characters
- `Right(string, n)` — Last n characters
- `Mid(string, start, length)` — Substring
- `Trim()` — Remove leading/trailing spaces

**Common Mistakes:**
- Mixing & with + (& is for strings, + is for numbers)
- Off-by-one errors with Left/Right/Mid
- Not handling empty strings

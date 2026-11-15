# VBA Exercise 1: Variables Basics — Solution

> Try the exercise in [README.md](README.md) first.

Ready-made files: `01_Variables_Basics_solution.xlsm` and `01_Variables_Basics_solution.bas`.

## Solution code

```vba
Sub DeclareVariables()
    Dim myAge As Integer
    Dim myName As String
    Dim myHeight As Double
    Dim isActive As Boolean
    
    myAge = 25
    myName = "John"
    myHeight = 5.9
    isActive = True
    
    MsgBox "Age: " & myAge & vbCrLf & _
           "Name: " & myName & vbCrLf & _
           "Height: " & myHeight & vbCrLf & _
           "Active: " & isActive
End Sub
```

## Explanation

**Key Concepts:**
- `Dim` keyword declares variables
- Each variable has a data type (Integer, String, Double, Boolean)
- Naming convention: camelCase or descriptive names
- VbCrLf adds line breaks in message boxes

**Common Mistakes:**
- Forgetting to declare variables (enable Option Explicit at module top)
- Using wrong data type (String for numbers, Integer for decimals)
- Not initializing values before using them

**Extension:**
```vba
' Add Option Explicit at the top of module to enforce declarations
Option Explicit

Sub VariablesWithScope()
    Dim localVar As Integer  ' Only available in this procedure
    
    localVar = 10
    Debug.Print "Local: " & localVar
End Sub
```

# VBA Exercise 51: Class Modules Basics — Solution

> Try the exercise in [README.md](README.md) first.

**Solution Code:**
```vba
' Class Module: Customer
Option Explicit

Private pCustomerID As Integer
Private pName As String
Private pEmail As String

Public Property Get CustomerID() As Integer
    CustomerID = pCustomerID
End Property

Public Property Let CustomerID(value As Integer)
    pCustomerID = value
End Property

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Sub Save()
    ' Save to database
    Debug.Print "Saving customer: " & pName
End Sub
```

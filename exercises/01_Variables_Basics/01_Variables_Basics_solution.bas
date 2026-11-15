Attribute VB_Name = "Exercise"
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

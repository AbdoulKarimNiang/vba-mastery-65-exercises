Attribute VB_Name = "Exercise"
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

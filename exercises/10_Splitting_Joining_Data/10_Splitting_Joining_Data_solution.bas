Attribute VB_Name = "Exercise"
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

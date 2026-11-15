Attribute VB_Name = "Exercise"
Sub FindReplace()
    Dim ws As Worksheet
    Dim foundCell As Range
    Dim replaceCount As Integer
    
    Set ws = ThisWorkbook.Sheets("Data")
    replaceCount = 0
    
    Set foundCell = ws.Columns(1).Find(What:="old", LookAt:=xlWhole)
    
    Do While Not foundCell Is Nothing
        foundCell.Value = "new"
        replaceCount = replaceCount + 1
        Set foundCell = ws.Columns(1).Find(What:="old", After:=foundCell)
    Loop
    
    MsgBox "Replaced " & replaceCount & " occurrences"
End Sub

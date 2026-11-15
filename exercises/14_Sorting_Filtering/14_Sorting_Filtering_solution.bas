Attribute VB_Name = "Exercise"
Sub SortAndFilter()
    Dim ws As Worksheet
    Dim dataRange As Range
    
    Set ws = ThisWorkbook.Sheets("Data")
    Set dataRange = ws.Range("A1:B11")
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B1"), Order:=xlDescending
        .SetRange dataRange
        .Header = xlYes
        .Apply
    End With
    
    dataRange.AutoFilter
    MsgBox "Sorted and filtered"
End Sub

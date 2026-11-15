Attribute VB_Name = "Exercise"
Sub NamedRanges()
    Dim ws As Worksheet
    Dim sumValue As Double
    Dim cell As Range
    
    Set ws = ThisWorkbook.Sheets("Data")
    
    ThisWorkbook.Names.Add Name:="SalesData", RefersTo:=ws.Range("A1:B10")
    
    sumValue = 0
    For Each cell In ws.Range("SalesData")
        If IsNumeric(cell.Value) Then
            sumValue = sumValue + cell.Value
        End If
    Next cell
    
    MsgBox "Total from named range: " & sumValue
End Sub

Function IsNumeric(val As Variant) As Boolean
    IsNumeric = Not IsError(val + 0)
End Function

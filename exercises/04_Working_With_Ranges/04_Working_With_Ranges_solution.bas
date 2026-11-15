Attribute VB_Name = "Exercise"
Sub WorkingWithRanges()
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim cell As Range
    Dim sumValue As Double
    Dim countValue As Long
    Dim avgValue As Double

    Set ws = ThisWorkbook.Worksheets("Data")
    Set dataRange = ws.Range("B2:F2")

    For Each cell In dataRange.Cells
        If Not IsEmpty(cell.Value) And IsNumeric(cell.Value) Then
            sumValue = sumValue + cell.Value
            countValue = countValue + 1
        End If
    Next cell

    If countValue > 0 Then avgValue = sumValue / countValue

    ws.Range("H1").Value = sumValue
    ws.Range("H2").Value = avgValue

    MsgBox "Sum: " & sumValue & vbCrLf & "Average: " & avgValue
End Sub

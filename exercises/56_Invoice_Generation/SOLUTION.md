# VBA Exercise 56: Invoice Generation — Solution

> Try the exercise in [README.md](README.md) first.

**Solution Code:**
```vba
Sub GenerateInvoice()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.Name = "Invoice_" & Format(Now, "yyyymmdd")
    
    ' Header
    ws.Range("A1").Value = "INVOICE"
    ws.Range("A3").Value = "Invoice Date:"
    ws.Range("B3").Value = Date
    
    ' Customer info
    ws.Range("A5").Value = "Bill To:"
    ws.Range("A6").Value = "Customer Name"
    ws.Range("A7").Value = "Address"
    
    ' Items table
    ws.Range("A10").Value = "Description"
    ws.Range("B10").Value = "Qty"
    ws.Range("C10").Value = "Unit Price"
    ws.Range("D10").Value = "Amount"
    
    ' Format as table
    With ws.Range("A10:D10").Font
        .Bold = True
        .Size = 12
    End With
    
    MsgBox "Invoice created: " & ws.Name
End Sub
```

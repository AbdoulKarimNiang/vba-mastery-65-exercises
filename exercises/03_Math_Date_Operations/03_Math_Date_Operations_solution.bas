Attribute VB_Name = "Exercise"
Sub MathDateOperations()
    Dim price As Double
    Dim tax As Double
    Dim totalPrice As Double
    Dim today As Date
    Dim endOfYear As Date
    Dim daysRemaining As Integer
    
    price = 100
    tax = 0.15
    totalPrice = price + (price * tax)
    
    today = Date
    endOfYear = DateSerial(Year(today), 12, 31)
    daysRemaining = endOfYear - today
    
    MsgBox "Price: $" & Format(price, "0.00") & vbCrLf & _
           "Total with tax: $" & Format(totalPrice, "0.00") & vbCrLf & _
           "Days until year end: " & daysRemaining
End Sub

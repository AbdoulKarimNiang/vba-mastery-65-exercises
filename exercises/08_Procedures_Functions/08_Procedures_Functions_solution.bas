Attribute VB_Name = "Exercise"
Sub CreateFunctions()
    Dim radius As Double
    Dim celsius As Double
    
    radius = 5
    celsius = 25
    
    MsgBox "Circle area: " & CircleArea(radius) & vbCrLf & _
           "25C in F: " & CelsiusToFahrenheit(celsius)
End Sub

Function CircleArea(r As Double) As Double
    CircleArea = 3.14159 * r * r
End Function

Function CelsiusToFahrenheit(c As Double) As Double
    CelsiusToFahrenheit = (c * 9/5) + 32
End Function

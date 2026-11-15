Attribute VB_Name = "Exercise"
Sub ManageWorksheets()
    Dim newSheet As Worksheet
    Dim originalSheet As Worksheet
    
    Set originalSheet = ThisWorkbook.ActiveSheet
    
    Set newSheet = ThisWorkbook.Sheets.Add
    newSheet.Name = "Report"
    newSheet.Range("A1").Value = "Hello"
    
    originalSheet.Activate
    MsgBox "Created new sheet: " & newSheet.Name
End Sub

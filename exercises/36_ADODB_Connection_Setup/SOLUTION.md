# VBA Exercise 36: ADODB Connection Setup — Solution

> Try the exercise in [README.md](README.md) first.

**Solution Code:**
```vba
Sub ADODBConnectionSetup()
    On Error GoTo ErrorHandler
    
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")
    
    Dim connStr As String
    connStr = "Provider=PostgreSQL OLE DB Provider;" & _
              "Server=localhost;" & _
              "Port=5432;" & _
              "Database=vba_training;" & _
              "User=vba_user;" & _
              "Password=vba_training_pwd;"
    
    conn.Open connStr
    MsgBox "Connected to PostgreSQL!"
    conn.Close
    
    Exit Sub
ErrorHandler:
    MsgBox "Connection failed: " & Err.Description
End Sub
```

# VBA Exercise 38: Inserting Data DB — Solution

> Try the exercise in [README.md](README.md) first.

**Solution Code:**
```vba
Sub InsertToDB()
    Dim conn As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=PostgreSQL OLE DB Provider;" & _
              "Server=localhost;Port=5432;" & _
              "Database=vba_training;User=vba_user;Password=vba_training_pwd;"
    
    ' Parameterized query (safer)
    Dim cmd As Object
    Set cmd = CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandText = "INSERT INTO customers (first_name, email) VALUES (?, ?)"
    cmd.Parameters(0).Value = "Jane Doe"
    cmd.Parameters(1).Value = "jane@example.com"
    cmd.Execute
    
    conn.Close
    MsgBox "Customer inserted"
End Sub
```

---

[Exercises 39-50 continue with database operations, transactions, stored procedures, etc.]

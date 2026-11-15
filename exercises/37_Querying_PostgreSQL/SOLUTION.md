# VBA Exercise 37: Querying PostgreSQL — Solution

> Try the exercise in [README.md](README.md) first.

**Solution Code:**
```vba
Sub QueryPostgres()
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=PostgreSQL OLE DB Provider;" & _
              "Server=localhost;Port=5432;" & _
              "Database=vba_training;User=vba_user;Password=vba_training_pwd;"
    
    sql = "SELECT product_id, name, price FROM products WHERE price > 50"
    Set rs = conn.Execute(sql)
    
    Dim row As Integer
    row = 1
    
    While Not rs.EOF
        With ThisWorkbook.ActiveSheet
            .Cells(row, 1).Value = rs.Fields("product_id").Value
            .Cells(row, 2).Value = rs.Fields("name").Value
            .Cells(row, 3).Value = rs.Fields("price").Value
        End With
        rs.MoveNext
        row = row + 1
    Wend
    
    rs.Close
    conn.Close
    MsgBox "Loaded " & (row - 1) & " products"
End Sub
```

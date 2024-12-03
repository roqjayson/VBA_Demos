' Enable activex data objects

Sub ReadData()

    Dim databaseFilename As String, connectionString As String
    
    databaseFilename = ThisWorkbook.Path & Application.PathSeparator & "Database1.accdb"
    
    ' Look for the right connectionString: https://www.connectionstrings.com/
    
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databaseFilename & ";Persist Security Info=False;"
    
    Dim conn As New ADODB.Connection
    conn.Open connectionString
    
    Dim rs As New ADODB.Recordset, query As String
    query = "SELECT * FROM salesdata"
    
    rs.Open query, conn
    
    With ActiveSheet
            .Cells.Clear
            .Range("A1").CopyFromRecordset rs
    End With
    
    conn.Close

End Sub

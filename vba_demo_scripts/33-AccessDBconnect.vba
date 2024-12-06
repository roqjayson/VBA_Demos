' Handle errors and close connection if errors occur

Sub ReadData()

    On Error GoTo ErrorHandler

    Dim databaseFilename As String, connectionString As String
    
    databaseFilename = ThisWorkbook.Path & Application.PathSeparator & "Databasess1.accdb"
    
    ' Look for the right connectionString: https://www.connectionstrings.com/
    
    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & databaseFilename & ";Persist Security Info=False;"
    
    Dim conn As New ADODB.Connection
    conn.Open connectionString
    
    Dim rs As New ADODB.Recordset, query As String
    query = "SELECT * FROM salesdata"
    
    rs.Open query, conn
    
cleanup:
    If Not (rs Is Nothing) Then
        If (rs.State And adStateOpen) = adStateOpen Then rs.Close
            Set rs = Nothing
        End If
    If Not (conn Is Nothing) Then
        If (conn.State And adStateOpen) = adStateOpen Then conn.Close
            Set conn = Nothing
        End If
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description
    GoTo cleanup


End Sub

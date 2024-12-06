' Three ways to scrape data
' Internet Explorer
' Chrome Selenium
' Microsoft XML Library


' Make sure to go to Tools > References > check the following
' Microsoft HTML Object Library
' Microsoft XML, v6.0 (or another available version like v3.0 or v4.0 for MSXML2.XMLHTTP60).

Private Sub GetHtmlFromUrl()

    Dim url As String
    url = "https://en.wikipedia.org/wiki/List_of_largest_companies_by_revenue"
    
    ' Declare the HTTP request and HTML document
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0") ' Late binding for compatibility

    Dim html As Object
    Set html = CreateObject("HTMLFile") ' Late binding for HTMLDocument
    
    ' Send HTTP request
    http.Open "GET", url, False
    http.Send
    
    ' Load HTML response into the HTML document
    html.body.innerHTML = http.responseText
    
    ' Read the tables in the webpage into a collection
    Dim coll As Collection
    Set coll = ReadTables(html)
    
    ' Write data to the worksheet
    WriteToSheet coll

End Sub

Function ReadTables(html As Object) As Collection

    Dim coll As New Collection
    Dim tables As Object
    Dim table As Object
    Dim row As Object
    Dim cell As Object
    Dim tableFound As Boolean
    
    ' Get all table elements
    Set tables = html.getElementsByTagName("table")
    
    For Each table In tables
        If InStr(table.innerText, "Rank") > 0 And InStr(table.innerText, "Revenue") > 0 Then
            tableFound = True
            Exit For
        End If
    Next table
    
    If tableFound Then
        For Each row In table.Rows
            Dim rowData As String
            rowData = ""
            For Each cell In row.Cells
                rowData = rowData & cell.innerText & vbTab
            Next cell
            coll.Add rowData
        Next row
    End If
    
    Set ReadTables = coll

End Function

Sub WriteToSheet(coll As Collection)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)
    ws.Cells.Clear ' Clear existing data
    
    Dim i As Long
    For i = 1 To coll.Count
        Dim rowData() As String
        rowData = Split(coll(i), vbTab)
        
        Dim j As Long
        For j = LBound(rowData) To UBound(rowData)
            ws.Cells(i, j + 1).Value = rowData(j)
        Next j
    Next i

End Sub


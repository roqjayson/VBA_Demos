' Make sure to build Sheet2

Public Sub UseQueryTable()

    Dim url As String
    url = "https://en.wikipedia.org/wiki/List_of_largest_companies_by_revenue"
    
    Dim table As QueryTable
    Set table = Sheet2.QueryTables.Add("URL;" & url, Sheet2.Range("A1"))
    
    With table
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "1"
        .WebFormatting = xlWebFormattingAll
        .Refresh
    End With
    


End Sub

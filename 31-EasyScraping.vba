' Issue with code is if we ran it again, it will insert the data to the right
Public Sub Main()
    
    Call ClearSheet
    Call UseQueryTable
    
    
End Sub

Private Sub ClearSheet()

    Dim ws As Worksheet
    For Each table In Sheet2.QueryTables
        table.Delete
    Next table
    
    Sheet2.Cells.Clear

End Sub



Public Sub UseQueryTable()

    Dim url As String
    url = "https://en.wikipedia.org/wiki/List_of_largest_companies_by_revenue"
    
    Dim table As QueryTable
    Set table = Sheet2.QueryTables.Add("URL;" & url, Sheet2.Range("A1"))
    
    With table
        .WebSelectionType = xlSpecifiedTables
        .WebTables = "1"
        .WebFormatting = xlWebFormattingNone
        .Refresh
    End With
    


End Sub

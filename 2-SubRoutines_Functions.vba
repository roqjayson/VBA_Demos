Sub ClearWorksheet()
    ActiveSheet.Cells.Clear
End Sub

Function PopulateSampleData() As Variant()
    
    ActiveSheet.Cells(2, 1).Value = "Product"
    ActiveSheet.Cells(2, 2).Value = "Category"
    ActiveSheet.Cells(2, 3).Value = "January Sales"
    ActiveSheet.Cells(2, 4).Value = "February Sales"

End Function


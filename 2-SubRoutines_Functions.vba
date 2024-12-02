Sub ClearWorksheet()
    ActiveSheet.Cells.Clear
End Sub

Function PopulateSampleData() As Variant()
    
    ActiveSheet.Cells(2, 1).Value = "Product"

End Function


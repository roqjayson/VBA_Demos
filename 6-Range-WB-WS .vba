Sub WorkbookWorksheetRangeDemo()

    ' Declare workbook and worksheet variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Set ws = ActiveSheet
    
    ' Change specific cell values individually
    ws.Cells(1, 1).value = "Header 1" ' Change cell A1
    ws.Cells(1, 2).value = "Header 2" ' Change cell B1
    ws.Cells(1, 3).value = "Header 3" ' Change cell C1

End Sub


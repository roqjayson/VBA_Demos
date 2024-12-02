Sub WorkbookWorksheetRangeDemo()

    ' Declare workbook and worksheet variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Set ws = ActiveSheet
    
    ' Select a range and work with it
    Set rng = ws.Range("A1:C5")
    
    ' Populate the range with data (fill A1 to C5 with some values)
    rng.value = "Hello"
End Sub


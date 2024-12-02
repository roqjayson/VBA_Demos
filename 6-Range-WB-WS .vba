Sub WorkbookWorksheetRangeDemo()

    ' Declare workbook and worksheet variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    
    ' Create a new workbook
    Set wb = Workbooks.Add
    
     ' Add a worksheet to the workbook
    Set ws = wb.Sheets.Add
    ws.Name = "DemoSheet"  ' Rename the sheet to "DemoSheet"
   
End Sub


Sub WorkbookWorksheetRangeDemo()

    ' Declare workbook and worksheet variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Set wb = ActiveWorkbook
    
    wb.SaveAs "C:\NewExcel\Book-Demo-Save.xlsx"

    MsgBox "Workbook is now saved!", vbInformation
    
    
    
End Sub


Sub WorkbookWorksheetRangeDemo()

    ' Declare workbook and worksheet variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range

    Set wb = ActiveWorkbook
    
    Set ws = ActiveSheet
    
    Application.DisplayAlerts = False
    
    wb.SaveAs "C:\NewExcel\Book-Demo-Save.xlsx"
    
    Application.DisplayAlerts = True
    
    MsgBox "Workbook is now saved!", vbInformation
    
    
    
End Sub


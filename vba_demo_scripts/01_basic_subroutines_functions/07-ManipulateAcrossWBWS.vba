Sub CopySRC_PasteTGT()

    Dim wbSource As Workbook, wbTarget As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim rngSource As Range, rngTarget As Range
    Dim lastRow As Long
    
    ' Open the source workbook
    Set wbSource = Workbooks.Open("C:\Users\roque\Desktop\Trainosys\Excel\VBA\Actual Demo\Source.xlsx")
    
    ' Set the source worksheet
    Set wsSource = wbSource.Sheets("Employees")
    
    ' Find the last row of data in the source worksheet (column A as an example)
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    ' Define the range to copy from the source worksheet
    Set rngSource = wsSource.Range("A1:C" & lastRow)
    
    ' Open the target workbook
    Set wbTarget = Workbooks.Open("C:\Users\roque\Desktop\Trainosys\Excel\VBA\Actual Demo\Target.xlsx")
    
    ' Set the target worksheet (for example, Sheet2)
    Set wsTarget = wbTarget.Sheets("Sheet1")
    
    ' Define the target range where the data will be pasted
    Set rngTarget = wsTarget.Range("A1")
    
    ' Copy data from the source workbook to the target workbook
    rngSource.Copy Destination:=rngTarget
    
    ' Example of manipulating data: Concatenate data in the target workbook (Column B and Column C)
    Dim i As Long
    For i = 1 To lastRow
        wsTarget.Cells(i, 4).Value = wsTarget.Cells(i, 2).Value & " - " & wsTarget.Cells(i, 3).Value
    Next i
    
    ' Example: Writing a value to a new cell in the source workbook
    wsSource.Cells(lastRow + 1, 1).Value = "Old Data sent to Target"
    
    ' Save and close the workbooks
    wbTarget.Save
    wbTarget.Close
    wbSource.Save
    wbSource.Close
    
    ' Notify the user
    MsgBox "Data manipulation completed across workbooks and worksheets!", vbInformation

End Sub


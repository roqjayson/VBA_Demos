Option Explicit

Private Sub UseForRangeCopy_ORIGINAL()

    Dim shData As Worksheet, shOutput As Worksheet
    Set shData = ThisWorkbook.Worksheets("Data")
    Set shOutput = ThisWorkbook.Worksheets("Output")
    
    shOutput.Range("A1").CurrentRegion.Offset(1).ClearContents
    
    Dim dTime As Double
    dTime = Timer
    
    Dim rg As Range
    Set rg = shData.Range("A1").CurrentRegion
    
    Dim i As Long, row As Long
    row = 2
    For i = 2 To rg.Rows.Count
    
        If rg.Cells(i, 9).Value = 10 Then

            rg.Rows(i).Copy
            shOutput.Range("A" & row).PasteSpecial xlPasteValues
            
            row = row + 1
            
        End If
        
    Next i
    
    Debug.Print "Time is: " & (Timer - dTime) * 1000
        


End Sub

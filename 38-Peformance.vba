Option Explicit

Private Sub UseForRangeCopy_ORIGINAL()

    Dim shData As Worksheet, shOutput As Worksheet
    Set shData = ThisWorkbook.Worksheets("Data")
    Set shOutput = ThisWorkbook.Worksheets("Output")
    
    shOutput.Range("A1").CurrentRegion.Offset(1).ClearContents
    
    Dim dTime As Double
    dTime = Timer
    
    Call TurnOffStuff
    
    Dim rg As Range
    Set rg = shData.Range("A1").CurrentRegion
    
    Dim i As Long, row As Long
    row = 2
    For i = 2 To rg.Rows.Count
    
        If rg.Cells(i, 9).Value = 10 Then

            shOutput.Range("A" & row).Resize(1, rg.Columns.Count).Value = rg.Rows(i).Value
            
            row = row + 1
            
        End If
        
    Next i
    
    Call TurnOnStuff
    
    Debug.Print "Time is: " & (Timer - dTime) * 1000
        


End Sub

Sub TurnOffStuff()

    Application.Calculation = xlCalculationManual
    
End Sub

Sub TurnOnStuff()

    Application.Calculation = xlCalculationAutomatic
    
End Sub

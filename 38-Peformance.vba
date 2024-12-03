Option Explicit

Private Sub UseForRangeCopy_ORIGINAL()

    Dim shData As Worksheet, shOutput As Worksheet
    Set shData = ThisWorkbook.Worksheets("Data")
    Set shOutput = ThisWorkbook.Worksheets("Output")
    
    shOutput.Range("A1").CurrentRegion.Offset(1).ClearContents
    
    Dim dTime As Double
    dTime = Timer
    
    Call TurnOffStuff
    
    Dim arr As Variant
    arr = shData.Range("A1").CurrentRegion.Value
    
    Dim i As Long, j As Long, row As Long
    row = 2
    For i = LBound(arr, 1) To UBound(arr, 1)
    
        If arr(i, 9) = 10 Then

            For j = LBound(arr, 2) To UBound(arr, 2)
                shOutput.Cells(row, j).Value = arr(i, j)
            
            Next j
            
            row = row + 1
            
        End If
        
    Next i
    
    Call TurnOnStuff
    
    Debug.Print "Time is: " & (Timer - dTime) * 1000
        


End Sub

Sub TurnOffStuff()

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
End Sub

Sub TurnOnStuff()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub

' You'll have to use Perf.xlsm to be able to utilize this script

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
    
    Dim criteriaRange As Range
    Set criteriaRange = ThisWorkbook.Worksheets("AdvFilter").Range("A1").CurrentRegion
    
    rg.AdvancedFilter xlFilterCopy, criteriaRange, shOutput.Range("A1:CH1")
    
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

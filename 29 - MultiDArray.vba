Option Explicit

Sub ReadingRange()

    Dim rg As Range
    Set rg = ActiveSheet.Range("A1").CurrentRegion
    
    Dim arr As Variant
    arr = rg.Value

End Sub

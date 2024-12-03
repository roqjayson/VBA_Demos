Option Explicit

Sub ReadingRange()
    
    Dim arr As Variant
    arr = ActiveSheet.Range("A1").CurrentRegion
    
    Dim i As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        arr(i, 2) = arr(i, 2) + 1
    Next i

End Sub 'Add Breakpoint here

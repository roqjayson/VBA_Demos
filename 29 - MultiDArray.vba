Option Explicit

Sub ReadingRange()
    
    Dim arr As Variant
    arr = ActiveSheet.Range("A1").CurrentRegion
    
    Dim i As Long
    For i = LBound(arr, 1) To UBound(arr, 1)
        Debug.Print arr(i, 1), arr(i, 3)
    Next i

End Sub

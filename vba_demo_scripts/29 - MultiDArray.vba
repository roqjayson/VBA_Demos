' Dynamic Size
' Delete previous output

Sub ReadingRange()
    
    Dim arr As Variant
    arr = ActiveSheet.Range("A1").CurrentRegion
    
    Dim i As Long
    For i = LBound(arr, 1) + 1 To UBound(arr, 1)
        arr(i, 2) = arr(i, 2) + 1
    Next i
    
    ActiveSheet.Range("E1").CurrentRegion.ClearContents
    
    Dim rowCount As Long, columnCount As Long
    rowCount = UBound(arr, 1)
    columnCount = UBound(arr, 2)
    ActiveSheet.Range("E1").Resize(rowCount, columnCount).Value = arr
    
    
End Sub

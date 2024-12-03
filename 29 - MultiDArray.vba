Option Explicit

Sub ReadingRange()

    Dim rg As Range
    Set rg = ActiveSheet.Range("A1").CurrentRegion
    
    Dim arr As Variant
    ReDim arr(1 To rg.Rows.Count, 1 To rg.Columns.Count)

    Dim i As Long, j As Long
    For i = 1 To rg.Rows.Count
        For j = 1 To rg.Columns.Count
            arr(i, j) = rg.Cells(i, j).Value
        Next j
    Next i

End Sub

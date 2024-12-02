Sub NestedLoops()
    Dim i As Integer, j As Integer
    ' Nested loops to print a multiplication table (1-5)
    For i = 1 To 5
        For j = 1 To 5
            MsgBox "Multiplication " & i & " x " & j & " = " & i * j
        Next j
    Next i
End Sub

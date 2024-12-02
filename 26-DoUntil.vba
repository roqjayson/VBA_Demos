Sub DoUntilLoop()
    Dim i As Integer
    i = 1
    ' Loop until i is greater than 5
    Do Until i > 5
        MsgBox "The value of i is " & i
        i = i + 1
    Loop
End Sub

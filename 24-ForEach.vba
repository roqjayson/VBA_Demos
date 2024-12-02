Sub ForEachNextLoop()
    Dim cell As Range
    ' Loop through all cells in range A1:A5
    For Each cell In Range("A1:A5")
        cell.Value = "Hello VBA"
    Next cell
End Sub

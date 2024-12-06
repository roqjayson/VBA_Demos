Sub DemonstrateArgumentPassing(ByVal valueParam As Integer, ByRef refParam As Integer)
    ' Modify parameters inside the subroutine
    
    ' For ByVal parameter: changes won't affect the original variable
    valueParam = valueParam * 2
    
    ' For ByRef parameter: changes will modify the original variable
    refParam = refParam * 2
End Sub

Sub CallDemonstrateArgumentPassing()
    ' Declare variables to pass to the subroutine
    Dim x As Integer
    Dim y As Integer
    
    ' Initialize variables
    x = 5
    y = 5
    
    ' Call the subroutine
    Call DemonstrateArgumentPassing(x, y)
    
    ' Display the results
    ' x will remain 5 (passed by value)
    ' y will be 10 (passed by reference)
    Worksheets("Sheet1").Range("A1").Value = "x: " & x
    Worksheets("Sheet1").Range("A2").Value = "y: " & y
End Sub


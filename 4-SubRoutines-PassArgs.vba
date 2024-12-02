Sub DemonstrateArgumentPassing(ByVal valueParam As Integer, ByRef refParam As Integer)
    ' Modify parameters inside the subroutine
    
    ' For ByVal parameter: changes won't affect the original variable
    valueParam = valueParam * 2
    
    ' For ByRef parameter: changes will modify the original variable
    refParam = refParam * 2
End Sub

Sub SimpleWorksheetOperation()
    ' Select a specific worksheet
    Worksheets("Sheet1").Activate
    
    ' Write values to different cells
    Range("A1").Value = "Hello VBA!"
    Range("B1").Value = Date
    Range("C1").Value = Time
    
End Sub

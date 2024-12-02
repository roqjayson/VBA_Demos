Sub SimpleWorksheetOperation()
    ' Select a specific worksheet
    Worksheets("Sheet1").Activate
    
    ' Write values to different cells
    Range("A1").Value = "Hello VBA!"
    Range("B1").Value = Date
    Range("C1").Value = Time

    ' Change the font color and formatting
    Range("A1").Font.Color = vbBlue
    Range("B1").Font.Bold = True
    Range("C1").Font.Italic = True
    
End Sub

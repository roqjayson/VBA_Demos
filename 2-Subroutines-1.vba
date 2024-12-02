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
     
    ' Autofit the columns
    Columns("A:C").AutoFit
    
    ' Add a simple calculation
    Range("A2").Value = 10
    Range("B2").Value = 5
    Range("C2").Value = "Sum:"
    Range("D2").Formula = "=A2+B2"
    
    
End Sub

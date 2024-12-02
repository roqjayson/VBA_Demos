' Procedure to write basic text and date information
Sub WriteBasicInformation(ws As Worksheet)
    With ws.Range("A1")
        .value = "Hello VBA!"
        .Font.Color = vbBlue
    End With
    
    With ws.Range("B1")
        .value = Date
        .Font.Bold = True
    End With
    
    With ws.Range("C1")
        .value = Time
        .Font.Italic = True
    End With
End Sub

' Procedure to perform a basic calculation
Sub PerformCalculation(ws As Worksheet)
    With ws
        .Range("A2").value = 10
        .Range("B2").value = 5
        .Range("C2").value = "Sum:"
        .Range("D2").Formula = "=A2+B2"
    End With
End Sub

' Procedure to style worksheet
Sub StyleWorksheet(ws As Worksheet)
    With ws
        .Columns("A:D").AutoFit
        .Range("A1:D2").Interior.Color = RGB(240, 240, 240)
    End With
End Sub

' Main Procedure to call all the subroutines
Sub ExecuteTasks(ws As Worksheet)
    ' Call the subroutines to perform tasks
    WriteBasicInformation ws
    PerformCalculation ws
    StyleWorksheet ws
End Sub

Sub RunAllTasks()
    ExecuteTasks Worksheets("Sheet1")
End Sub

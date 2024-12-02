Sub MultipleErrorHandlers()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double
    
    ' Setting up error handling for runtime errors
    On Error GoTo ErrorHandler

    ' Initializing variables
    x = 10
    y = 0 ' This will cause a divide-by-zero error

    ' First operation: Division by zero
    result = x / y ' This triggers a divide-by-zero error
    Debug.Print "Result: " & result

    Exit Sub ' Ensure we skip the error handler unless an error occurs

ErrorHandler:
    If Err.Number = 11 Then ' Division by zero error
        Debug.Print "Error: Division by zero occurred!"
    Else
        Debug.Print "Error: " & Err.Description
    End If
    MsgBox "Error: " & Err.Description, vbCritical

End Sub


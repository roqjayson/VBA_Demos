Sub ResetErrorHandling()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double
    
    ' Set up error handling
    On Error GoTo ErrorHandler

    ' Initializing variables
    x = 10
    y = 0 ' This will cause a divide-by-zero error
    
    ' First operation: Division by zero
    result = x / y ' This triggers a divide-by-zero error
    Debug.Print "Result: " & result

    ' Disable error handling
    On Error GoTo 0

    ' This will now trigger a runtime error because error handling was turned off
    result = x / y ' This will now show a runtime error again
    
    Exit Sub

ErrorHandler:
    ' Error handling code
    Debug.Print "An error occurred: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical

End Sub


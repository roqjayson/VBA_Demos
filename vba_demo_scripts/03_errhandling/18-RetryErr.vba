Sub RetryAfterError()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double

    ' Initializing variables
    x = 10
    y = 0 ' This will cause a divide-by-zero error

    ' Set up error handling
    On Error GoTo ErrorHandler

    ' First operation: Division by zero
    result = x / y ' This triggers a divide-by-zero error
    Debug.Print "Result: " & result

    Exit Sub

ErrorHandler:
    ' Error handling code
    Debug.Print "Error: Division by zero occurred, trying again..."

    ' Retry logic - we change 'y' to a non-zero value and retry
    y = 5
    Resume ' Retry the operation with updated 'y'

End Sub


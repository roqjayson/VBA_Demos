
Part 1: Basic Error Handling using On Error GoTo
In this part, we demonstrate the most basic error handling using On Error GoTo to direct the flow of execution to an error handling section when an error occurs.

Sub BasicErrorHandling()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double

    ' Initializing variables
    x = 10
    y = 0 ' This will cause a divide-by-zero error

    ' Basic error handling setup
    On Error GoTo ErrorHandler

    ' Attempting division by zero
    result = x / y ' This will trigger an error

    ' Normal code execution will be skipped if error occurs
    Debug.Print "Result: " & result

    Exit Sub ' Ensure we don't reach the error handler unless an error occurs

ErrorHandler:
    ' Error handling code
    Debug.Print "An error occurred: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical

End Sub

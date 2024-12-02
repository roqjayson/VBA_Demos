Sub ClearErrorDemo()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double

    ' Initializing variables
    x = 10
    y = 0 ' This will cause a divide-by-zero error

    ' Set up error handling
    On Error Resume Next

    ' First operation: Division by zero
    result = x / y ' This triggers a divide-by-zero error
    Debug.Print "Result: " & result

    ' Clear the error object
    Err.Clear

    ' This operation will not raise an error because the previous one was cleared
    result = x + y
    Debug.Print "New result after clearing error: " & result

    Exit Sub

ErrorHandler:
    ' Error handling code
    Debug.Print "An error occurred: " & Err.Description
    MsgBox "Error: " & Err.Description, vbCritical

End Sub


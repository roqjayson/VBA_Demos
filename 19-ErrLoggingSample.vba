Sub SampleProcedureWithErrorLogging()

    On Error GoTo ErrorHandler ' Enable error handling

    ' Example code with an intentional error (divide by zero)
    Dim x As Integer
    Dim y As Integer
    Dim result As Double

    x = 10
    y = 0 ' This will cause a divide by zero error

    result = x / y ' This will raise a runtime error

    Exit Sub ' Ensures the error handler is skipped if no error occurs

ErrorHandler:
    ' Call the LogError procedure to log the error details
    LogError "SampleProcedureWithErrorLogging", Err.Number, Err.Description

    ' Optionally, show an error message
    MsgBox "An error has occurred. Please check the log file for details.", vbCritical

    ' Resume normal code execution after logging the error
    Resume Next

End Sub

Sub LogError(ByVal procName As String, ByVal errorNumber As Long, ByVal errorDescription As String)
    ' This subroutine writes error details to a log file

    Dim logFilePath As String
    Dim logFile As Integer
    Dim logMessage As String
    Dim currentDate As String

    ' Define the log file path (in the same directory as the workbook)
    logFilePath = ThisWorkbook.Path & "\ErrorLog.txt"

    ' Get the current date and time for the log entry
    currentDate = Now

    ' Create the log message
    logMessage = "Date: " & currentDate & vbCrLf & _
                 "Procedure: " & procName & vbCrLf & _
                 "Error Number: " & errorNumber & vbCrLf & _
                 "Error Description: " & errorDescription & vbCrLf & _
                 "----------------------------------------" & vbCrLf

    ' Open the log file for appending (create it if it doesn't exist)
    logFile = FreeFile
    Open logFilePath For Append As logFile

    ' Write the log message to the file
    Print #logFile, logMessage

    ' Close the log file
    Close logFile
End Sub



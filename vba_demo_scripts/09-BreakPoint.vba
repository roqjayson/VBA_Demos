Sub BreakpointDemo()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double

    ' Initializing variables
    x = 10
    y = 5

    ' This line is where you should manually set a breakpoint (Click in the left margin or press F9)
    Debug.Print "BreakpointDemo started."

    ' Performing a simple calculation
    result = x + y
    Debug.Print "Result after addition (x + y): " & result ' Expected: 15

    ' Add a breakpoint here to pause the execution and inspect values before continuing
    Debug.Print "BreakpointDemo completed."
    
End Sub


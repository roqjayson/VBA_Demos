Sub AssertDemo()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double

    ' Initializing variables
    x = 10
    y = 5

    ' Assert that x should be 10 (this will pass)
    Debug.Assert x = 10

    ' Assert that y should be 0 (this will fail, causing an alert)
    Debug.Assert y = 0 ' This line will trigger an assertion failure

    ' Performing a simple calculation
    result = x + y
    Debug.Print "Sum of x and y: " & result

End Sub



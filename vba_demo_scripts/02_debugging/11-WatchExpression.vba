Sub WatchExpressionDemo()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double

    ' Initializing variables
    x = 10
    y = 5

    ' Perform some calculations (Intentional mistake here)
    result = x + y
    Debug.Print "Sum of x and y: " & result
    
    result = x * y
    Debug.Print "Product of x and y: " & result
    
    ' Watch expression for variable "x"
    ' Right-click on variable "x" or "y" and select "Add Watch" to monitor its value in the Watch Window.
    ' You can also use the Immediate Window to check its current value during execution.
    
End Sub

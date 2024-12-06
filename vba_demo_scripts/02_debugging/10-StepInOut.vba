Sub StepThroughDemo()

    ' Variables declaration
    Dim x As Integer
    Dim y As Integer
    Dim result As Double

    ' Initializing variables
    x = 10
    y = 5

    ' Step-by-step execution
    Debug.Print "Step-by-step execution begins."

    ' Step 1: Addition
    result = x + y
    Debug.Print "Step 1 - Addition result: " & result ' Expected: 15

    ' Step 2: Multiplication
    result = x * y
    Debug.Print "Step 2 - Multiplication result: " & result ' Expected: 50

    ' Step 3: Subtraction
    result = x - y
    Debug.Print "Step 3 - Subtraction result: " & result ' Expected: 5

    ' Step 4: Division (Intentional mistake here: Division by zero later)
    result = x / y
    Debug.Print "Step 4 - Division result: " & result ' Expected: 2
    
    ' Fix the division by zero error
    ' result = x / 0 ' Uncomment to cause an error

    Debug.Print "Step-by-step execution ends."

End Sub

' Module1

Sub Main()

    Dim account As New clsAccount
    
    Call account.CreditAccount(100)
    Call account.DebitAccount(25)
    
    Debug.Print account.balance

    

End Sub

' Let's say the requirements has changed
' We are now required to add fees on top of the credit or debit

' modAccount

Sub CreditAccount(balance As Currency, fees As Double, amount As Currency)
    balance = balance + (amount * fees)
End Sub

Sub DebitAccount(balance As Currency, fees As Double, amount As Currency)
    balance = balance - (amount * fees)
End Sub

' Create a new module here called modAccount

' clsAccount

Option Explicit

Public balance As Currency
Public fees As Double

Public Sub CreditAccount(amount As Currency)
    balance = (balance + amount) + (amount * fees)
End Sub

Public Sub DebitAccount(amount As Currency)
    balance = (balance - amount) - (amount * fees)
End Sub

' Click the General at the top and change to class

Private Sub Class_Initialize()
    balance = 50
    fees = 0.05
End Sub

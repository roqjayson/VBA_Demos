Sub Main()

    Dim balance As Currency
    balance = 0
    

    Call CreditAccount(balance, fees, 100)
    Call DebitAccount(balance, fees, 25)
    
    Debug.Print balance

    

End Sub

' Let's say the requirements has changed
' We are now required to add fees on top of the credit or debit



Sub CreditAccount(balance As Currency, fees As Double, amount As Currency)
    balance = balance + (amount * fees)
End Sub

Sub DebitAccount(balance As Currency, fees As Double, amount As Currency)
    balance = balance - (amount * fees)
End Sub

' Create a new module here called modAccount

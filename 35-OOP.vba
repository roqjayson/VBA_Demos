Sub Main()

    Dim balance As Currency
    balance = 0
    
    ' credit
    Call CreditAccount(balance, 100)
    
    ' debit
    Call DebitAccount(balance, 25)
    
    'Common issues include
    'Easy to type the static numbers wrong
    'Difficult to update for nonVBA user
    'Code is not clear
    
    

End Sub


Sub CreditAccount(balance As Currency, amount As Currency)
    balance = balance + amount
End Sub

Sub DebitAccount(balance As Currency, amount As Currency)
    balance = balance - amount
End Sub

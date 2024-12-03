Sub Main()

    Dim account As New clsAccount
    
    Call account.CreditAccount(100)
    Call account.DebitAccount(25)
    account.Balance = 999
    Debug.Print account.Balance

    

End Sub

' Let's say the requirements has changed
' We are now required to add fees on top of the credit or debit


Sub CreditAccount(Balance As Currency, fees As Double, amount As Currency)
    Balance = Balance + (amount * fees)
End Sub

Sub DebitAccount(Balance As Currency, fees As Double, amount As Currency)
    Balance = Balance - (amount * fees)
End Sub

' Create a new module here called modAccount


Option Explicit

Private m_balance As Currency
Private m_fees As Double

Public Property Get Balance() As Currency
    Balance = m_balance
End Property

' Comment this out

'Public Property Let Balance(ByVal newBalance As Currency)
 '   m_balance = newBalance
'End Property

Public Sub CreditAccount(amount As Currency)
    Balance = (m_balance + amount) + (amount * fees)
End Sub

Public Sub DebitAccount(amount As Currency)
    Balance = (m_balance - amount) - (amount * fees)
End Sub

' Click the General at the top and change to class

Private Sub Class_Initialize()
    m_balance = 50
    fees = 0.05
End Sub





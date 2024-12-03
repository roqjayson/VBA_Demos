
Private Sub UserForm_Initialize()

    Me.Caption = "Data Entry Form"
    Me.Width = 500
    Me.Height = 500
    
    ComboBox1.AddItem "Add"
    ComboBox1.AddItem "Update"
    ComboBox1.AddItem "Delete"
    ComboBox1.ListIndex = 0
    
    Label1.Caption = "Name"
    Label2.Caption = "Age"
    Label3.Caption = "Email"
    
    CommandButton1.Caption = "Submit"
    CommandButton2.Caption = "Cancel"
    
    PopulateListBox
    
End Sub
Private Sub PopulateListBox()

    ListBox1.Clear
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim i As Long
    i = 2
    Do Until ws.Cells(i, 1).Value = ""
        ListBox1.AddItem ws.Cells(i, 1).Value & "-" & ws.Cells(i, 2).Value & "-" & ws.Cells(i, 3).Value
        i = i + 1
    Loop
End Sub





Private Sub OperationComboBox_Change()

    If ComboBox1.Value = "Add" Then
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox1.Value = ""
        TextBox2.Value = ""
        TextBox3.Value = ""
    ElseIf ComboBox1.Value = "Update" Then
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        ListBox1.Visible = True
        PopulateListBox
    ElseIf ComboBox1.Value = "Delete" Then
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ListBox1.Visible = True
        PopulateListBox
    End If
End Sub


Private Sub CommandButton1_Click()

    If Not ValidateInputs() Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    
    If ComboBox1.Value = "Add" Then
        Dim nextRow As Long
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
        ws.Cells(nextRow, 1).Value = TextBox1.Value
        ws.Cells(nextRow, 2).Value = TextBox2.Value
        ws.Cells(nextRow, 3).Value = TextBox3.Value
        MsgBox "Data succesfully added!", vbInformation
    End If
    
    If ComboBox1.Value = "Update" Then
        Dim selectedRow As Long
        selectedRow = ListBox1.ListIndex + 2
        
        If selectedRow > 1 Then
            ws.Cells(selectedRow, 1).Value = TextBox1.Value
            ws.Cells(selectedRow, 2).Value = TextBox2.Value
            ws.Cells(selectedRow, 3).Value = TextBox3.Value
            MsgBox "Data succesfully updated!", vbInformation
        Else
             MsgBox "Please select a record to update!", vbExclamation
        
        End If
    End If
    
    If ComboBox1.Value = "Delete" Then
        Dim selectedRowDelete As Long
        selectedRowDelete = ListBox1.ListIndex + 2
        
        If selectedRowDelete > 1 Then
           ws.Rows(selectedRowDelete).Delete
           MsgBox "Data succesfully deleted!", vbInformation
        Else
             MsgBox "Please select a record to delete!", vbExclamation
        
        End If
    End If
    
    
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""
    
    PopulateListBox
    
End Sub

Private Function ValidateInputs() As Boolean

    If Trim(TextBox1.Value) = "" Then
        MsgBox "Name cannot be empty!", vbExclamation
        TextBox1.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    If Trim(TextBox1.Value) = "" Or Not IsNumeric(TextBox2.Value) Then
        MsgBox "Age must be a valid number!", vbExclamation
        TextBox2.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    If Trim(TextBox3.Value) = "" Or Not IsValidEmail(TextBox3.Value) Then
        MsgBox "Please enter a valid email address!", vbExclamation
        TextBox3.SetFocus
        ValidateInputs = False
        Exit Function
    End If
    
    
    
    ValidateInputs = True


End Function

Private Function IsValidEmail(ByVal email As String) As Boolean
    
    If email Like "*@*.*" Then
        IsValidEmail = True
    Else
        IsValidEmail = False
    End If

End Function



Private Sub CommandButton2_Click()

    Me.Hide
    
End Sub

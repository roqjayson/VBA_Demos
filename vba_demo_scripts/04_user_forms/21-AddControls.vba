Private Sub UserForm_Initialize()
    ' Set default properties for the form
    Me.Caption = "CRUD Data Entry Form"
    Me.Width = 600
    Me.Height = 400

    ' Populate the ComboBox with operation choices
    ComboBox1.AddItem "Add"
    ComboBox1.AddItem "Update"
    ComboBox1.AddItem "Delete"
    ComboBox1.ListIndex = 0 ' Set default value to "Add"
    
    ' Set labels' text
    Label1.Caption = "Name"
    Label2.Caption = "Age"
    Label3.Caption = "Email"
    
    ' Set command button texts
    CommandButton1.Caption = "Submit"
    CommandButton2.Caption = "Cancel"

    ' Populate ListBox with existing records from the "Data" sheet
    PopulateListBox
End Sub

Private Sub OperationComboBox_Change()
    ' Enable/Disable controls based on selected operation
    If ComboBox1.Value = "Add" Then
        ' For Add operation, enable TextBoxes and clear them
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox1.Value = ""
        TextBox2.Value = ""
        TextBox3.Value = ""
        ListBox1.Visible = False
    ElseIf ComboBox1.Value = "Update" Then
        ' For Update operation, show the ListBox and enable TextBoxes
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        ListBox1.Visible = True
        PopulateListBox
    ElseIf ComboBox1.Value = "Delete" Then
        ' For Delete operation, show the ListBox and disable TextBoxes
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ListBox1.Visible = True
        PopulateListBox
    End If
End Sub

Private Sub CommandButton1_Click()
    ' Perform action based on selected operation
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data") ' Ensure "Data" sheet exists

    ' Add new data
    If ComboBox1.Value = "Add" Then
        Dim nextRow As Long
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(nextRow, 1).Value = TextBox1.Value
        ws.Cells(nextRow, 2).Value = TextBox2.Value
        ws.Cells(nextRow, 3).Value = TextBox3.Value
        MsgBox "Data successfully added!", vbInformation
    End If

    ' Update data
    If ComboBox1.Value = "Update" Then
        Dim selectedRow As Long
        selectedRow = ListBox1.ListIndex + 2 ' Adjust for row index (starting from 2)
        
        If selectedRow > 1 Then ' Ensure a record is selected
            ws.Cells(selectedRow, 1).Value = TextBox1.Value
            ws.Cells(selectedRow, 2).Value = TextBox2.Value
            ws.Cells(selectedRow, 3).Value = TextBox3.Value
            MsgBox "Data successfully updated!", vbInformation
        Else
            MsgBox "Please select a record to update!", vbExclamation
        End If
    End If

    ' Delete data
    If ComboBox1.Value = "Delete" Then
        Dim selectedRowDelete As Long
        selectedRowDelete = ListBox1.ListIndex + 2 ' Adjust for row index (starting from 2)
        
        If selectedRowDelete > 1 Then ' Ensure a record is selected
            ws.Rows(selectedRowDelete).Delete
            MsgBox "Data successfully deleted!", vbInformation
        Else
            MsgBox "Please select a record to delete!", vbExclamation
        End If
    End If

    ' Clear the textboxes after submitting
    TextBox1.Value = ""
    TextBox2.Value = ""
    TextBox3.Value = ""

    ' Refresh the ListBox
    PopulateListBox
End Sub

Private Sub CommandButton2_Click()
    ' Close the form without saving
    Me.Hide
End Sub

Private Sub RecordListBox_Click()
    ' When a record is clicked in the ListBox, display the data in the textboxes for update or delete
    If ListBox1.ListIndex > -1 Then
        Dim ws As Worksheet
        Set ws = ThisWorkbook.Sheets("Data")
        
        ' Get the selected record's row
        Dim selectedRow As Long
        selectedRow = RecordListBox.ListIndex + 2 ' Adjust for row index (starting from 2)
        
        ' Populate textboxes with selected record data
        TextBox1.Value = ws.Cells(selectedRow, 1).Value
        TextBox2.Value = ws.Cells(selectedRow, 2).Value
        TextBox3.Value = ws.Cells(selectedRow, 3).Value
    End If
End Sub

Private Sub PopulateListBox()
    ' Populate ListBox with data from the "Data" sheet
    ListBox1.Clear
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    Dim i As Long
    i = 2 ' Start from row 2 (assuming row 1 has headers)
    Do Until ws.Cells(i, 1).Value = ""
        ListBox1.AddItem ws.Cells(i, 1).Value & " - " & ws.Cells(i, 2).Value & " - " & ws.Cells(i, 3).Value
        i = i + 1
    Loop
End Sub


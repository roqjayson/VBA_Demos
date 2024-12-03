' Enable Microsoft WinHTTP
' Add scripting runtime for dict as well
' https://github.com/VBA-tools/VBA-JSON
Sub ReadFormAPI()

    Dim request As New WinHttpRequest
    
    request.Open "Get", "https://api.nationalize.io/?name=charles"
    request.Send
    
    If request.Status <> 200 Then
        MsgBox request.responseText
        Exit Sub
    End If
    
    Dim response As Object
    Set response = JsonConverter.ParseJson(request.responseText)
    
    
    Debug.Print response("name")
    
    Dim countries As Collection
    Set countries = response("country")
    
    Dim country As Dictionary
    For Each country In countries
        Debug.Print country("country_id"), FormatPercent(country("probability"))
    Next country
    


End Sub

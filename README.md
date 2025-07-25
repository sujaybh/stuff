# stuff


regex.Pattern = "^GFCMA\.2025\.Q[1-4]\.S\d+\s+\(\d{1,2}/\d{1,2}\s+\d{1,2}/\d{1,2}\)$"

Function GetCurrentSprintId(boardId As String, bearerToken As String) As String
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim json As Object
    Dim sprints As Object
    Dim sprint As Object
    Dim sprintName As String
    Dim i As Integer
    
    ' Create HTTP request object
    Set http = CreateObject("MSXML2.XMLHTTP")
    
    ' Build URL with active state filter
    url = "https://horizon.jira2.bankofamerica.com/rest/agile/1.0/board/" & boardId & "/sprint?state=active"
    
    ' Configure request
    http.Open "GET", url, False
    http.setRequestHeader "Authorization", "Bearer " & bearerToken
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Accept", "application/json"
    http.setRequestHeader "X-Atlassian-Token", "no-check"
    
    ' Send request
    On Error GoTo ErrorHandler
    http.send
    
    ' Check if request was successful
    If http.Status <> 200 Then
        GetCurrentSprintId = "Error: HTTP " & http.Status & " - " & http.statusText
        Exit Function
    End If
    
    ' Parse JSON response
    response = http.responseText
    Set json = JsonConverter.ParseJson(response)
    Set sprints = json("values")
    
    ' Loop through active sprints and find matching pattern
    For i = 0 To sprints.Count - 1
        Set sprint = sprints(i)
        sprintName = sprint("name")
        
        ' Check if sprint name matches the required format
        If IsValidSprintFormat(sprintName) Then
            GetCurrentSprintId = CStr(sprint("id"))
            Exit Function
        End If
    Next i
    
    GetCurrentSprintId = "No matching active sprint found"
    Exit Function
    
ErrorHandler:
    GetCurrentSprintId = "Error: " & Err.Description
End Function

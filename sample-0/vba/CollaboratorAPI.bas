Attribute VB_Name = "CollaboratorAPI"
Option Explicit

Public keyJWT As String
Const URL = "https://<your-domen>"

Function fetchHttpRequest(cURL As String, requestBody As String, _
                    ByRef objHTTPStatus, Optional jwtToken) As String
    Dim objHTTP As Object 'New WinHttp.WinHttpRequest
    
    On Error GoTo LineErr
    
    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    objHTTP.Open "POST", cURL, False
    If Not IsMissing(jwtToken) Then
        objHTTP.SetRequestHeader "Authorization", "Bearer " & jwtToken
    End If
    objHTTP.SetRequestHeader "Content-Type", "application/json;charset=UTF-8"
    objHTTP.Send (requestBody)
    objHTTPStatus = objHTTP.Status
    fetchHttpRequest = objHTTP.ResponseText
Exit Function

LineErr:
    Debug.Print Err.Number
End Function


Sub getAccessToken(login, pass)
    Dim cURL As String
    Dim requestBody As String
    Dim objHTTPStatus
    Dim jsonString As String
    Dim res As Object
    
    cURL = URL & "/api/rest.php/auth/session"
    requestBody = "{""email"":""" & login & """,""password"":""" & pass & """}"
    
    jsonString = fetchHttpRequest(cURL, requestBody, objHTTPStatus)
    
    Select Case objHTTPStatus
        Case 200
         Call InitScriptEngine
         Set res = DecodeJsonString(jsonString)
         keyJWT = res.jwt_token
         Debug.Print "User: " & res.firstname & " " & res.secondname & vbNewLine & keyJWT
        
        Case 400
            keyJWT = ""
            Debug.Print "Ошибка доступа: " & objHTTPStatus & vbNewLine & "Пользователь не найден. Проверьте правильность ввода логина и пароля."
         
        Case Else
            keyJWT = ""
            Debug.Print "Ошибка доступа: " & objHTTPStatus & vbNewLine & "Проверьте, все ли параметры запроса правильные."
            
    End Select
End Sub

Sub test()
  Call getAccessToken("<your-login>", "<your-pass>")
  'Печатаем токен в текст активного документа
  Selection.TypeText Text:="keyJWT:" & vbNewLine & keyJWT
End Sub

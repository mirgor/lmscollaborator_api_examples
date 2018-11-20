Attribute VB_Name = "modParser"
Option Explicit
' Need added a reference to
'    Microsoft ScriptControl

Private ScriptEngine As Object ' ScriptControl

Function CreateObjectx86(Optional sProgID, Optional bClose = False)

    Static oWnd As Object
    Dim bRunning As Boolean

    #If Win64 Then
        bRunning = InStr(TypeName(oWnd), "HTMLWindow") > 0
        If bClose Then
            If bRunning Then oWnd.Close
            Exit Function
        End If
        If Not bRunning Then
            Set oWnd = CreateWindow()
            oWnd.execScript "Function CreateObjectx86(sProgID): Set CreateObjectx86 = CreateObject(sProgID): End Function", "VBScript"
        End If
        Set CreateObjectx86 = oWnd.CreateObjectx86(sProgID)
    #Else
        Set CreateObjectx86 = CreateObject(sProgID)
    #End If

End Function
Function CreateWindow()

    ' source http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
    Dim sSignature, oShellWnd, oProc

    On Error Resume Next
    sSignature = Left(CreateObject("Scriptlet.TypeLib").GUID, 38)
    CreateObject("WScript.Shell").Run "%systemroot%\syswow64\mshta.exe about:""about:<head><script>moveTo(-32000,-32000);document.title='x86Host'</script><hta:application showintaskbar=no /><object id='shell' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shell.putproperty('" & sSignature & "',document.parentWindow);</script></head>""", 0, False
    Do
        For Each oShellWnd In CreateObject("Shell.Application").Windows
            Set CreateWindow = oShellWnd.GetProperty(sSignature)
            If Err.Number = 0 Then Exit Function
            Err.Clear
        Next
    Loop

End Function

Public Sub InitScriptEngine()
    Set ScriptEngine = CreateObjectx86("MSScriptControl.ScriptControl") 'New ScriptControl
    ScriptEngine.Language = "JScript"
    ScriptEngine.Timeout = 10000000
    ScriptEngine.AddCode "function getProperty(jsonObj, propertyName) { return jsonObj[propertyName]; } "
    ScriptEngine.AddCode "function getKeys(jsonObj) { var keys = new Array(); for (var i in jsonObj) { keys.push(i); } return keys; } "
End Sub

Public Function DecodeJsonString(ByVal jsonString As String)
    Set DecodeJsonString = ScriptEngine.Eval("(" + jsonString + ")")
End Function

Public Function GetProperty(ByVal jsonObject As Object, ByVal propertyName As String) 'As Variant
    GetProperty = ScriptEngine.Run("getProperty", jsonObject, propertyName)
End Function

Public Function GetObjectProperty(ByVal jsonObject As Object, ByVal propertyName As String) 'As Object
    Set GetObjectProperty = ScriptEngine.Run("getProperty", jsonObject, propertyName)
End Function

Public Function GetKeys(ByVal jsonObject As Object) As String()
    Dim Length As Integer
    Dim KeysArray() As String
    Dim KeysObject As Object
    Dim Index As Integer
    Dim Key As Variant

    Set KeysObject = ScriptEngine.Run("getKeys", jsonObject)
    Length = GetProperty(KeysObject, "length")
    ReDim KeysArray(Length - 1)
    Index = 0
    For Each Key In KeysObject
        KeysArray(Index) = Key
        Index = Index + 1
    Next
    GetKeys = KeysArray
End Function

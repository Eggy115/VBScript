' List Internet Explorer Connection Settings


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_ConnectionSettings")

For Each strIESetting in colIESettings
    Wscript.Echo "Allow Internet programs: " & _
        strIESetting.AllowInternetPrograms
    Wscript.Echo "Autoconfiguration URL: " & strIESetting.AutoConfigURL
    Wscript.Echo "Auto disconnect: " & strIESetting.AutoDisconnect
    Wscript.Echo "Autoconfiguration proxy detection mode: " & _
        strIESetting.AutoProxyDetectMode
    Wscript.Echo "Data encryption: " & strIESetting.DataEncryption
    Wscript.Echo "Default: " & strIESetting.Default
    Wscript.Echo "Default gateway: " & strIESetting.DefaultGateway
    Wscript.Echo "Dialup server: " & strIESetting.DialUpServer
    Wscript.Echo "Disconnect idle time: " & strIESetting.DisconnectIdleTime
    Wscript.Echo "Encrypted password: " & strIESetting.EncryptedPassword
    Wscript.Echo "IP address: " & strIESetting.IPAddress
    Wscript.Echo "IP header compression: " & _
        strIESetting.IPHeaderCompression
    Wscript.Echo "Modem: " & strIESetting.Modem
    Wscript.Echo "Name: " & strIESetting.Name
    Wscript.Echo "Network logon: " & strIESetting.NetworkLogon
    Wscript.Echo "Network protocols: " & strIESetting.NetworkProtocols
    Wscript.Echo "Primary DNS server: " & strIESetting.PrimaryDNS
    Wscript.Echo "Primary WINS server: " & strIESetting.PrimaryWINS
    Wscript.Echo "Proxy: " & strIESetting.Proxy
    Wscript.Echo "Proxy override: " & strIESetting.ProxyOverride
    Wscript.Echo "Proxy server: " & strIESetting.ProxyServer
    Wscript.Echo "Redial attempts: " & strIESetting.RedialAttempts
    Wscript.Echo "Redial wait: " & strIESetting.RedialWait
    Wscript.Echo "Script fileame: " & strIESetting.ScriptFileName
    Wscript.Echo "Secondary DNS server: " & strIESetting.SecondaryDNS
    Wscript.Echo "Secondary WINS server: " & strIESetting.SecondaryWINS
    Wscript.Echo "Server assigned IP address: " & _
        strIESetting.ServerAssignedIPAddress
    Wscript.Echo "Server assigned name server: " & _
        strIESetting.ServerAssignedNameServer
    Wscript.Echo "Software compression: " & strIESetting.SoftwareCompression
Next

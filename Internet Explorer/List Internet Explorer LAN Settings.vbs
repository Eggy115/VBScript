' List Internet Explorer LAN Settings


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_LANSettings")

For Each strIESetting in colIESettings
    Wscript.Echo "Autoconfiguration proxy: " & strIESetting.AutoConfigProxy
    Wscript.Echo "Autoconfiguration URL: " & strIESetting.AutoConfigURL
    Wscript.Echo "Autoconfiguration Proxy detection mode: " & _
        strIESetting.AutoProxyDetectMode
    Wscript.Echo "Proxy: " & strIESetting.Proxy
    Wscript.Echo "Proxy override: " & strIESetting.ProxyOverride
    Wscript.Echo "Proxy server: " & strIESetting.ProxyServer
Next

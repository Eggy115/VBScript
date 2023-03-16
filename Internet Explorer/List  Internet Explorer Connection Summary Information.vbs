' List  Internet Explorer Connection Summary Information


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_ConnectionSummary")

For Each strIESetting in colIESettings
    Wscript.Echo "Connection preference: " & _
        strIESetting.ConnectionPreference
    Wscript.Echo "HTTP 1.1. enabled: " & strIESetting.EnableHTTP11
    Wscript.Echo "Proxy HTTP 1.1. enabled: " & strIESetting.ProxyHTTP11
Next

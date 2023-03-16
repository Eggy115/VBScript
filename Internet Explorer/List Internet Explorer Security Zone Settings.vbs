' List Internet Explorer Security Zone Settings


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_Security")

For Each strIESetting in colIESettings
    Wscript.Echo "Zone name: " & strIESetting.Zone
    Wscript.Echo "Security level: " & strIESetting.Level
Next

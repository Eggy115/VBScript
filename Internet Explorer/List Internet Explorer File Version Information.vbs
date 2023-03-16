' List Internet Explorer File Version Information


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_FileVersion")

For Each strIESetting in colIESettings
    Wscript.Echo "Company: " & strIESetting.Company
    Wscript.Echo "Date: " & strIESetting.Date
    Wscript.Echo "File name: " & strIESetting.File
    Wscript.Echo "Path: " & strIESetting.Path
    Wscript.Echo "File size: " & strIESetting.Size
    Wscript.Echo "Version: " & strIESetting.Version
Next

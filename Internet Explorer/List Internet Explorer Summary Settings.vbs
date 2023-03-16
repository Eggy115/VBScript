' List Internet Explorer Summary Settings


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_Summary")

For Each strIESetting in colIESettings
    Wscript.Echo "Active printer: " & strIESetting.ActivePrinter
    Wscript.Echo "Build: " & strIESetting.Build
    Wscript.Echo "Cipher strength: " & strIESetting.CipherStrength
    Wscript.Echo "Content advisor: " & strIESetting.ContentAdvisor
    Wscript.Echo "IE Administration Kit installed: " & _
        strIESetting.IEAKInstall
    Wscript.Echo "Language: " & strIESetting.Language
    Wscript.Echo "Name: " & strIESetting.Name
    Wscript.Echo "Path: " & strIESetting.Path
    Wscript.Echo "Product ID: " & strIESetting.ProductID
    Wscript.Echo "Version: " & strIESetting.Version
Next

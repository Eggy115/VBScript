' List Internet Explorer Cache Settings


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_Cache")

For Each strIESetting in colIESettings
    Wscript.Echo "Page refresh type: " & strIESetting.PageRefreshType
    Wscript.Echo "Temporary Internet files folder: " & _
        strIESetting.TempInternetFilesFolder
Next

' List Internet Explorer COM Object Settings


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\root\cimv2\Applications\MicrosoftIE")

Set colIESettings = objWMIService.ExecQuery _
    ("Select * from MicrosoftIE_Object")

For Each strIESetting in colIESettings
    Wscript.Echo "Code base: " & strIESetting.CodeBase
    Wscript.Echo "Program file: " & strIESetting.ProgramFile
    Wscript.Echo "Status: " & strIESetting.Status
Next

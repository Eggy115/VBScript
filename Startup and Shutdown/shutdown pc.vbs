

' Shut Down a Computer


strComputer = "."
Set objWMIService = GetObject_
    ("winmgmts:{impersonationLevel=impersonate,(Shutdown)}\\" & _
        strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
 
For Each objOperatingSystem in colOperatingSystems
    objOperatingSystem.Win32Shutdown(1)
Next

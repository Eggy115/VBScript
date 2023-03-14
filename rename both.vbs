' Rename a Computer and Computer Account


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colComputers = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")

For Each objComputer in colComputers
    err = objComputer.Rename("WebServer")
Next

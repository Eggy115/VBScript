' Modify System Startup Delay


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colStartupCommands = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")

For Each objStartupCommand in colStartupCommands
    objStartupCommand.SystemStartupDelay = 10
    objStartupCommand.Put_
Next


' Create a System Restore Point


CONST DEVICE_DRIVER_INSTALL = 10
CONST BEGIN_SYSTEM_CHANGE = 100

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestore")
errResults = objItem.CreateRestorePoint _
    ("Scripted restore", DEVICE_DRIVER_INSTALL, BEGIN_SYSTEM_CHANGE)

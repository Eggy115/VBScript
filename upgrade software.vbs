' Upgrade Software


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSoftware = objWMIService.ExecQuery _
    ("Select * from Win32__Product Where Name = 'Personnel Database'")

For Each objSoftware in colSoftware
    errReturn = objSoftware.Upgrade("c:\scripts\database2.msi")
Next

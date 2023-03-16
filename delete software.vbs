' Delete Software


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSoftware = objWMIService.ExecQuery _
    ("Select * from Win32_Product Where Name = 'Personnel database'")

For Each objSoftware in colSoftware
    objSoftware.Uninstall()
Next

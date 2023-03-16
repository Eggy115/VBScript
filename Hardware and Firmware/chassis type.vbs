' Identifying Computer Chassis Type


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colChassis = objWMIService.ExecQuery _
    ("Select * from Win32_SystemEnclosure")

For Each objChassis in colChassis
    For i = Lbound(objChassis.ChassisTypes) to Ubound(objChassis.ChassisTypes)
        Wscript.Echo objChassis.ChassisTypes(i)
    Next
Next

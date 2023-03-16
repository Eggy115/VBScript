' Verify That SQL Server is Installed on a Computer


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colServices = objWMIService.ExecQuery _
    ("Select * from Win32_Service Where Name = 'MSSQLServer'")

If colServices.Count > 0 Then
    For Each objService in colServices
        Wscript.Echo "SQL Server is " & objService.State & "."
    Next
Else
    Wscript.Echo "SQL Server is not installed on this computer."
End If

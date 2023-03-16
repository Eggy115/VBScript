' Verify Computer Availability


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colPingedComputers = objWMIService.ExecQuery _
    ("Select * from Win32_PingStatus Where Address = '192.168.1.37'")

For Each objComputer in colPingedComputers
    If objComputer.StatusCode = 0 Then
        Wscript.Echo "Remote computer responded."
    Else
        Wscript.Echo "Remote computer did not respond."
   End If
Next

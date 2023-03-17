' List Specific Events from an Event Log


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
        ("Select * from Win32_NTLogEvent Where Logfile = 'System' and " _
            & "EventCode = '6008'")

Wscript.Echo "Improper shutdowns: " & colLoggedEvents.Count

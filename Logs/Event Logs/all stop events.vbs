
' List All Stop Events


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'System'" _
        & " and SourceName = 'SaveDump'")

For Each objEvent in colLoggedEvents
    Wscript.Echo "Event date: " & objEvent.TimeGenerated
    Wscript.Echo "Description: " & objEvent.Message
Next


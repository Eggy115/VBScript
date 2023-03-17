' List Events From a Specific Event Log


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = 'Application'")

For Each objEvent in colLoggedEvents
    Wscript.Echo "Category: " & objEvent.Category
    Wscript.Echo "Computer Name: " & objEvent.ComputerName
    Wscript.Echo "Event Code: " & objEvent.EventCode
    Wscript.Echo "Message: " & objEvent.Message
    Wscript.Echo "Record Number: " & objEvent.RecordNumber
    Wscript.Echo "Source Name: " & objEvent.SourceName
    Wscript.Echo "Time Written: " & objEvent.TimeWritten
    Wscript.Echo "Event Type: " & objEvent.Type
    Wscript.Echo "User: " & objEvent.User
Next

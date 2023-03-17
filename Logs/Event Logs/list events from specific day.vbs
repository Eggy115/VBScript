

' List Events For a Specific Day From An Event Log


Const CONVERT_TO_LOCAL_TIME = True

Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")
DateToCheck = CDate("2/18/2002")
dtmStartDate.SetVarDate DateToCheck, CONVERT_TO_LOCAL_TIME
dtmEndDate.SetVarDate DateToCheck + 1, CONVERT_TO_LOCAL_TIME

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where TimeWritten >= '" _ 
        & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'") 

For Each objEvent in colEvents
    Wscript.Echo "Category: " & objEvent.Category
    Wscript.Echo "Computer Name: " & objEvent.ComputerName
    Wscript.Echo "Event Code: " & objEvent.EventCode
    Wscript.Echo "Message: " & objEvent.Message
    Wscript.Echo "Record Number: " & objEvent.RecordNumber
    Wscript.Echo "Source Name: " & objEvent.SourceName
    Wscript.Echo "Time Written: " & objEvent.TimeWritten
    Wscript.Echo "Event Type: " & objEvent.Type
    Wscript.Echo "User: " & objEvent.User
    Wscript.Echo objEvent.LogFile
Next

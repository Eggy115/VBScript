' Copy the Previous Day's Event Log Events to a Database


Set objConn = CreateObject("ADODB.Connection")
Set objRS = CreateObject("ADODB.Recordset")

objConn.Open "DSN=EventLogs;"
objRS.CursorLocation = 3
objRS.Open "SELECT * FROM EventTable" , objConn, 3, 3

Set dtmStartDate = CreateObject("WbemScripting.SWbemDateTime")
Set dtmEndDate = CreateObject("WbemScripting.SWbemDateTime")

DateToCheck = Date - 1
dtmEndDate.SetVarDate Date, True
dtmStartDate.SetVarDate DateToCheck, True

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where TimeWritten >= '" _ 
        & dtmStartDate & "' and TimeWritten < '" & dtmEndDate & "'") 

For each objEvent in colEvents
    objRS.AddNew
    objRS("Category") = objEvent.Category
    objRS("ComputerName") = objEvent.ComputerName
    objRS("EventCode") = objEvent.EventCode
    objRS("Message") = objEvent.Message
    objRS("RecordNumber") = objEvent.RecordNumber
    objRS("SourceName") = objEvent.SourceName
    objRS("TimeWritten") = objEvent.TimeWritten
    objRS("Type") = objEvent.Type
    objRS("User") = objEvent.User
    objRS.Update
Next

objRS.Close
objConn.Close

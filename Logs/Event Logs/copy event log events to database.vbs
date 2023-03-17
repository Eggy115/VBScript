' Copy Event Log Events to a Database


Set objConn = CreateObject("ADODB.Connection")
Set objRS = CreateObject("ADODB.Recordset")

objConn.Open "DSN=EventLogs;"
objRS.CursorLocation = 3
objRS.Open "SELECT * FROM EventTable" , objConn, 3, 3
strComputer = "."

Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colRetrievedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent")

For Each objEvent in colRetrievedEvents
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

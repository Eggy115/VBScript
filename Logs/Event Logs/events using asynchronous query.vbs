


' List Events Using an Asynchronous  Query


Const POPUP_DURATION = 10
Const OK_BUTTON = 0

Set objWSHShell = Wscript.CreateObject("Wscript.Shell")

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objSink = WScript.CreateObject("WbemScripting.SWbemSink","SINK_")
objWMIService.InstancesOfAsync objSink, "Win32_NTLogEvent"
Error = objWshShell.Popup("Starting event retrieval", POPUP_DURATION, _
    "Event Retrieval", OK_BUTTON)

Sub SINK_OnCompleted(iHResult, objErrorObject, objAsyncContext)
    WScript.Echo "Asynchronous operation is done."
End Sub

Sub SINK_OnObjectReady(objEvent, objAsyncContext)
    Wscript.Echo "Category: " & objEvent.Category
    Wscript.Echo "Computer Name: " & objEvent.ComputerName
    Wscript.Echo "Event Code: " & objEvent.EventCode
    Wscript.Echo "Message: " & objEvent.Message
    Wscript.Echo "Record Number: " & objEvent.RecordNumber
    Wscript.Echo "Source Name: " & objEvent.SourceName
    Wscript.Echo "Time Written: " & objEvent.TimeWritten
    Wscript.Echo "Event Type: " & objEvent.Type
    Wscript.Echo "User: " & objEvent.User
End Sub


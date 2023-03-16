' List Audited Internet Explorer Events from the Security Log


On Error Resume Next

strComputer = "."
Set dtmDate = CreateObject("WbemScripting.SWbemDateTime")
Set objWMIService = GetObject("winmgmts:" _
    & "{(Security)}!\\" & strComputer & "\root\cimv2")
Set colLoggedEvents = objWMIService.ExecQuery _
        ("SELECT * FROM Win32_NTLogEvent WHERE Logfile = 'Security' AND " _
            & "EventCode = '560'")

For Each objEvent in colLoggedEvents
    errResult = _
        InStr(objEvent.Message,"\REGISTRY\MACHINE\SOFTWARE\Microsoft\") 
    If errResult <> 0 Then
        Select Case objEvent.EventType
            Case 4 strEventType = "Success"
            Case 5 strEventType = "Failure"
        End Select
        Wscript.Echo objEvent.User
        dtmDate.Value = objEvent.TimeWritten
        dtmTimeWritten = dtmDate.GetVarDate
        Wscript.Echo "Time written: " & dtmTimeWritten
        Wscript.Echo strEventType
        Wscript.Echo "Record number: " & objEvent.RecordNumber & VbCrLf
        Wscript.Echo objEvent.Message
        Wscript.Echo 
    End If
Next

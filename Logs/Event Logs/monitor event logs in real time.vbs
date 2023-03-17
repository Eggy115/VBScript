
' Monitor Event Logs in Real Time


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate, (Security)}!\\" & _
        strComputer & "\root\cimv2")

Set colMonitoredEvents = objWMIService.ExecNotificationQuery _    
    ("Select * from __instancecreationevent where " _
        & "TargetInstance isa 'Win32_NTLogEvent' " _
            & "and TargetInstance.EventCode = '533' ")

Do
    Set objLatestEvent = colMonitoredEvents.NextEvent
     strAlertToSend = objLatestEvent.TargetInstance.User _ 
         & " attempted to access DatabaseServer."
     Wscript.Echo strAlertToSend
Loop

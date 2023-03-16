' Monitor Changes in Computer Power Status


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colMonitoredEvents = objWMIService.ExecNotificationQuery _
    ("Select * from Win32_PowerManagementEvent")

Do
    Set strLatestEvent = colMonitoredEvents.NextEvent
    Wscript.Echo strLatestEvent.EventType
Loop

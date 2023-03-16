' Monitor Processor Use


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.Swbemrefresher")
Set objProcessor = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfOS_Processor").objectSet
intThresholdViolations = 0
objRefresher.Refresh

Do
    For Each intProcessorUse in objProcessor
        If intProcessorUse.PercentProcessorTime > 90 Then
            intThresholdViolations = intThresholdViolations + 1
                If intThresholdViolations = 10 Then
                    intThresholdViolations = 0
                    Wscript.Echo "Processor usage threshold exceeded."
                End If
        Else
            intThresholdViolations = 0
        End If
    Next
    Wscript.Sleep 6000
    objRefresher.Refresh
Loop

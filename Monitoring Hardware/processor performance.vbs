' Monitor Processor Performance


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfOS_Processor").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "C1 Transitions Per Second: " & _
            objItem.C1TransitionsPersec
        Wscript.Echo "C2 Transitions Per Second: " & _
            objItem.C2TransitionsPersec
        Wscript.Echo "C3 Transitions Per Second: " & _
            objItem.C3TransitionsPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "DPC Rate: " & objItem.DPCRate
        Wscript.Echo "DPCs Queued Per Second: " & objItem.DPCsQueuedPersec
        Wscript.Echo "Interrupts Per Second: " & objItem.InterruptsPersec
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Percent C1 Time: " & objItem.PercentC1Time
        Wscript.Echo "Percent C2 Time: " & objItem.PercentC2Time
        Wscript.Echo "Percent C3 Time: " & objItem.PercentC3Time
        Wscript.Echo "Percent DPC Time: " & objItem.PercentDPCTime
        Wscript.Echo "Percent Idle Time: " & objItem.PercentIdleTime
        Wscript.Echo "Percent Interrupt Time: " & objItem.PercentInterruptTime
        Wscript.Echo "Percent Privileged Time: " & _
            objItem.PercentPrivilegedTime
        Wscript.Echo "Percent Processor Time: " & objItem.PercentProcessorTime
        Wscript.Echo "Percent User Time: " & objItem.PercentUserTime
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

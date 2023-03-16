
' Monitor Available Memory


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set objMemory = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfOS_Memory").objectSet
objRefresher.Refresh

Do
    For Each intAvailableBytes in objMemory
        If intAvailableBytes.AvailableMBytes < 4 Then
            Wscript.Echo "Available memory has fallen below 4 megabytes."
        End If
    Next
    objRefresher.Refresh
Loop

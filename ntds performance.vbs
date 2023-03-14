' Monitor NTDS Performance


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_NTDS_NTDS").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
    Wscript.Echo "Directory service threads in use: " & _
        objItem.DSThreadsInUse
    Wscript.Sleep 2000
    objRefresher.Refresh
    Next
Next

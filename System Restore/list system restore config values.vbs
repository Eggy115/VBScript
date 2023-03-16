' List System Restore Configuration Values


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set colItems = objWMIService.ExecQuery("Select * from SystemRestoreConfig")

For Each objItem in colItems
    Wscript.Echo "Disk Percent: " & objItem.DiskPercent
    Wscript.Echo "Global Interval (in seconds): " & objItem.RPGlobalInterval 
    Wscript.Echo "Life Interval (in seconds): " & objItem.RPLifeInterval
    If objItem.RPSessionInterval = 0 Then
        Wscript.Echo "Session Interval: Feature not enabled." 
    Else
        Wscript.Echo "Session Interval (in seconds): " & _
            objItem.RPSessionInterval
    End If
Next

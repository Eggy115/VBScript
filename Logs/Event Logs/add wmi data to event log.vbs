' Add WMI Data to an Event Log Entry



Const EVENT_FAILED = 2

Set objShell = Wscript.CreateObject("Wscript.Shell")
Set objNetwork = Wscript.CreateObject("Wscript.Network")
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDiskDrives = objWMIService.ExecQuery _
    ("Select * from win32_perfformatteddata_perfdisk_logicaldisk")

For Each objDisk in colDiskDrives
    strDriveSpace = objDisk.Name & " " & objDisk.FreeMegabytes _
        & VbCrLf
Next

strEventDescription = "Payroll application could not be installed on " _ 
    & objNetwork.UserDomain & "\" & objNetwork.ComputerName _ 
        & " by user " & objNetwork.UserName & _
            ". Free space on each drive is: " & strDriveSpace
objShell.LogEvent EVENT_FAILED, strEventDescription

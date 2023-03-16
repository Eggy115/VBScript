' List Recovery Configuration Options


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colRecoveryOptions = objWMIService.ExecQuery _
    ("Select * from Win32_OSRecoveryConfiguration")

For Each objOption in colRecoveryOptions 
    Wscript.Echo "Auto reboot: " & objOption.AutoReboot
    Wscript.Echo "Debug File Path: " & objOption.DebugFilePath
    Wscript.Echo "Debug Info Type: " & objOption.DebugInfoType
    Wscript.Echo "Kernel Dump Only: " & objOption.KernelDumpOnly
    Wscript.Echo "Name: " & objOption.Name
    Wscript.Echo "Overwrite Existing Debug File: " & _
        objOption.OverwriteExistingDebugFile
    Wscript.Echo "Send Administrative Alert: " & objOption.SendAdminAlert
    Wscript.Echo "Write Debug Information: " & objOption.WriteDebugInfo
    Wscript.Echo "Write to System Log: " & objOption.WriteToSystemLog
Next

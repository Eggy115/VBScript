' Modify Recovery Configuration Options


Const COMPLETE_MEMORY_DUMP = 1

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colRecoveryOptions = objWMIService.ExecQuery _
    ("Select * from Win32_OSRecoveryConfiguration")

For Each objOption in colRecoveryOptions 
    objOption.DebugInfoType = COMPLETE_MEMORY_DUMP
    objOption.DebugFilePath = "c:\scripts\memory.dmp"
    objOption.OverWriteExistingDebugFile = False
    objOption.Put_
Next


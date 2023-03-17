
' List System Event Log Properties


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='System'")

For Each objLogFile in colLogFiles
    Wscript.Echo objLogFile.NumberOfRecords
Next

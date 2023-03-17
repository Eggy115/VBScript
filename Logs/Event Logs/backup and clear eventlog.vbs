' Back Up and Clear an Event Log


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Backup)}!\\" & _
        strComputer & "\root\cimv2")

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='Application'")

For Each objLogfile in colLogFiles
    errBackupLog = objLogFile.BackupEventLog("c:\scripts\application.evt")
    If errBackupLog <> 0 Then        
        Wscript.Echo "The Application event log could not be backed up."
    Else
        objLogFile.ClearEventLog()
    End If
Next

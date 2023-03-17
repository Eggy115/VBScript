' Create Unique File Names for Event Log Backups


dtmThisDay = Day(Date)
dtmThisMonth = Month(Date)
dtmThisYear = Year(Date)
strBackupName = dtmThisYear & "_" & dtmThisMonth & "_" & dtmThisDay

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Backup)}!\\" & _
        strComputer & "\root\cimv2")

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile where LogFileName='Application'")

For Each objLogfile in colLogFiles
    objLogFile.BackupEventLog("c:\scripts\" & strBackupName & _
        "_application.evt")
    objLogFile.ClearEventLog()
Next

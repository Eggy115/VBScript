
' Back Up and Clear Large Event Logs


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate, (Backup, Security)}!\\" _
        & strComputer & "\root\cimv2")

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile")

For Each objLogfile in colLogFiles
    If objLogFile.FileSize > 100000 Then
       strBackupLog = objLogFile.BackupEventLog _
           ("c:\scripts\" & objLogFile.LogFileName & ".evt")
       objLogFile.ClearEventLog()
    End If
Next

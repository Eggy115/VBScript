
' List Event Log Properties


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set objInstalledLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile")

For each objLogfile in objInstalledLogFiles
    Wscript.Echo "Name: " &  objLogfile.LogFileName 
    Wscript.Echo "Maximum Size: " &  objLogfile.MaxFileSize 
    If objLogfile.OverWriteOutdated > 365 Then
        Wscript.Echo "Overwrite Outdated Records: Never." 
    ElseIf objLogfile.OverWriteOutdated = 0 Then
        Wscript.Echo "Overwrite Outdated Records: As needed." 
    Else
        Wscript.Echo "Overwrite Outdated Records After: " &  _
            objLogfile.OverWriteOutdated & " days" 
    End If
Next

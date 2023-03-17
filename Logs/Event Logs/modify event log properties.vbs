' Modify Event Log Properties


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate,(Security)}!\\" & _
        strComputer & "\root\cimv2")

Set colLogFiles = objWMIService.ExecQuery _
    ("Select * from Win32_NTEventLogFile")

For each objLogfile in colLogFiles
    strLogFileName = objLogfile.Name
    Set wmiSWbemObject = GetObject _
        ("winmgmts:{impersonationLevel=Impersonate}!\\.\root\cimv2:" _
            & "Win32_NTEventlogFile.Name='" & strLogFileName & "'")
    wmiSWbemObject.MaxFileSize = 2500000000
    wmiSWbemObject.OverwriteOutdated = 14
    wmiSWbemObject.Put_ 
Next

' Monitor Domain Controller Performance


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDatabases = objWMIService.ExecQuery _
    ("Select * from Win32_PerfFormattedData_NTDS_NTDS")

For Each objADDatabase in colDatabases
    Wscript.Echo "DS threads in use: " & objADDatabase.DSThreadsInUse
    Wscript.Echo "LDAP bind time: " & objADDatabase.LDAPBindTime
    Wscript.Echo "LDAP client sessions: " & objADDatabase.LDAPClientSessions
Next

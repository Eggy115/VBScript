
' Monitor Active Directory Database Performance


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colDatabases = objWMIService.ExecQuery _
    ("Select * from Win32_PerfFormattedData_Esent_Database " _
        & "Where Name = 'NT Directory'")

For Each objADDatabase in colDatabases
    Wscript.Echo "Database cache hit percent: " & _
        objADDatabase.DatabaseCachePercentHit
Next

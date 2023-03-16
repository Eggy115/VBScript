
' List Logon Session Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_LogonSession")

For Each objItem in colItems
    Wscript.Echo "Authentication Package: " & objItem.AuthenticationPackage
    Wscript.Echo "Logon ID: " & objItem.LogonId
    Wscript.Echo "Logon Type: " & objItem.LogonType
    Wscript.Echo "Start Time: " & objItem.StartTime
    Wscript.Echo
Next

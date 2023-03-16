' Enumerating Resultant Set of Policy System Services


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_SystemService")

For Each objItem in colItems
    Wscript.Echo "Service: " & objItem.Service
    Wscript.Echo "Precedence: " & objItem.Precedence
    Wscript.Echo "SDDL String: " & objItem.SDDLString
    Wscript.Echo "Startup Mode: " & objItem.StartupMode
    Wscript.Echo
Next

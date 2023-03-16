' List Environment Variables on a Computer


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Environment")

For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "System Variable: " & objItem.SystemVariable
    Wscript.Echo "User Name: " & objItem.UserName
    Wscript.Echo "Variable Value: " & objItem.VariableValue
Next

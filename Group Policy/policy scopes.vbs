' Enumerating Resultant Set of Policy Scopes of Management


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_SOM")

For Each objItem in colItems
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "Blocked: " & objItem.Blocked
    Wscript.Echo "Blocking: " & objItem.Blocking
    Wscript.Echo "Reason: " & objItem.Reason
    Wscript.Echo "SOM Order: " & objItem.SOMOrder
    Wscript.Echo "Type: " & objItem.Type
    Wscript.Echo
Next

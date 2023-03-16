' List Resultant Set of Policy Group Policy Links


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_GPLink")

For Each objItem in colItems
    Wscript.Echo "GPO: " & objItem.GPO
    Wscript.Echo "Applied Order: " & objItem.AppliedOrder
    Wscript.Echo "Enabled: " & objItem.Enabled
    Wscript.Echo "Link Order: " & objItem.LinkOrder
    Wscript.Echo "No Overrride: " & objItem.NoOverride
    Wscript.Echo "SOM Order: " & objItem.SOMOrder
    Wscript.Echo
Next


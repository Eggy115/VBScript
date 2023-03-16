' List Resultant Set of Policy Policy Settings


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_PolicySetting")

For Each objItem in colItems
    Wscript.Echo "GPO ID: " & objItem.GPOID
    Wscript.Echo "ID: " & objItem.ID
    Wscript.Echo "Precedence: " & objItem.Precedence
    Wscript.Echo "SOM ID: " & objItem.SOMID
    Wscript.Echo
Next

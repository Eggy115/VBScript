
' List Resultant Set of Policy Audit Policies


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_AuditPolicy")

For Each objItem in colItems  
    Wscript.Echo "Category: " & objItem.Category
    Wscript.Echo "Precedence: " & objItem.Precedence
    Wscript.Echo "Failure: " & objItem.Failure
    Wscript.Echo "Success: " & objItem.Success
    Wscript.Echo
Next


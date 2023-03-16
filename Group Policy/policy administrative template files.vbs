' List Resultant Set of Policy Administrative Template Files


Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
 
strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery _
    ("Select * from RSOP_AdministrativeTemplateFile")

For Each objItem in colItems  
    Wscript.Echo "GPO ID: " & objItem.GPOID
    dtmConvertedDate.Value = objItem.LastWriteTime
    dtmCreationTime = dtmConvertedDate.GetVarDate
    Wscript.Echo "Last Write Time: " & dtmCreationTime 
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo
Next


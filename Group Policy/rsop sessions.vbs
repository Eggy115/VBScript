' List RSOP Sessions


Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
 
strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery("Select * from RSOP_Session")

For Each objItem in colItems
    Wscript.Echo "ID: " & objItem.ID
    dtmConvertedDate.Value = objItem.CreationTime
    dtmTime = dtmConvertedDate.GetVarDate
    Wscript.Echo "Creation Time: " & dtmTime 
    Wscript.Echo "Flags: " & objItem.Flags
    For Each strSecurityGroup in objItem.SecurityGroups
        Wscript.Echo "Security group: " & strSecurityGroup
    Next
    Wscript.Echo "Site: " & objItem.Site
    Wscript.Echo "Slow Link: " & objItem.SlowLink
    Wscript.Echo "SOM: " & objItem.SOM
    Wscript.Echo "Target Name: " & objItem.TargetName
    Wscript.Echo "TTL (Minutes): " & objItem.TTLMinutes
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo
Next

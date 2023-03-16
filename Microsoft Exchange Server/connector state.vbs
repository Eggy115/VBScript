' List Exchange Connector State Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" &  _
        strComputer & "\root\cimv2\Applications\Exchange")

Set colItems = objWMIService.ExecQuery _
    ("Select * from ExchangeConnectorState")

For Each objItem in colItems
    Wscript.Echo "Distinguished name: " & objItem.DN
    Wscript.Echo "Group distinguished name: " & objItem.GroupDN
    Wscript.Echo "Group GUID: " & objItem.GroupGUID
    Wscript.Echo "GUID: " & objItem.GUID
    Wscript.Echo "Is up: " & objItem.IsUp
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo
Next

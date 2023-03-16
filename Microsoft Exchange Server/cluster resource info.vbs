' List Exchange Cluster Resource Information



On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" &  _
        strComputer & "\root\CIMV2\Applications\Exchange")

Set colItems = objWMIService.ExecQuery _
     ("Select * from ExchangeClusterResource")

For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Owner: " & objItem.Owner
    Wscript.Echo "State: " & objItem.State
    Wscript.Echo "Type: " & objItem.Type
    Wscript.Echo "Virtual machine: " & objItem.VirtualMachine
    Wscript.Echo
Next

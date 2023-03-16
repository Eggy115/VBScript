' List Exchange Virtual Server Queue Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_QueueVirtualServer")

For Each objItem in colItems
    Wscript.Echo "Global actions supported: " & _
        objItem.GlobalActionsSupported
    Wscript.Echo "Global stop: " & objItem.GlobalStop
    Wscript.Echo "Protocol name: " & objItem.ProtocolName
    Wscript.Echo "Virtual machine: " & objItem.VirtualMachine
    Wscript.Echo "Virtual server name: " & _
        objItem.VirtualServerName
    Wscript.Echo
Next

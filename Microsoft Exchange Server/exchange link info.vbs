' List Exchange Link Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_Link")

For Each objItem in colItems
    Wscript.Echo "Action freeze: " & objItem.ActionFreeze
    Wscript.Echo "Action kick: " & objItem.ActionKick
    Wscript.Echo "Action thaw: " & objItem.ActionThaw
    Wscript.Echo "Extended state info: " & _
        objItem.ExtendedStateInfo
    Wscript.Echo "Global stop: " & objItem.GlobalStop
    Wscript.Echo "Link distinguished name: " & objItem.LinkDN
    Wscript.Echo "Link ID: " & objItem.LinkID
    Wscript.Echo "Link name: " & objItem.LinkName
    Wscript.Echo "Message count: " & objItem.MessageCount
    Wscript.Echo "Next scheduled connection: " & _
        objItem.NextScheduledConnection
    Wscript.Echo "Oldest message: " & objItem.OldestMessage
    Wscript.Echo "Protocol name: " & objItem.ProtocolName
    Wscript.Echo "Size: " & objItem.Size
    Wscript.Echo "State action: " & objItem.StateActive
    Wscript.Echo "State flags: " & objItem.StateFlags
    Wscript.Echo "State frozen: " & objItem.StateFrozen
    Wscript.Echo "State ready: " & objItem.StateReady
    Wscript.Echo "State remote: " & objItem.StateRemote
    Wscript.Echo "State retry: " & objItem.StateRetry
    Wscript.Echo "State scheduled: " & objItem.StateScheduled
    Wscript.Echo "Support link actions: " & _
        objItem.SupportLinkActions
    Wscript.Echo "Type currently unreachable: " & _
        objItem.TypeCurrentlyUnreachable
    Wscript.Echo "Type deferred deilvery: " & _
        objItem.TypeDeferredDelivery
    Wscript.Echo "Type internal: " & objItem.TypeInternal
    Wscript.Echo "Type local delivery: " & _
        objItem.TypeLocalDelivery
    Wscript.Echo "Type pending categorization: " & _
        objItem.TypePendingCategorization
    Wscript.Echo "Type pending routing: " & _
        objItem.TypePendingRouting
    Wscript.Echo "Type pending submission: " & _
        objItem.TypePendingSubmission
    Wscript.Echo "Type remote delivery: " & _
        objItem.TypeRemoteDelivery
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "Virtual machine: " & objItem.VirtualMachine
    Wscript.Echo "Virtual server name: " & _
        objItem.VirtualServerName
    Wscript.Echo
Next

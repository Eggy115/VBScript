' List Exchange Queued Message Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_QueuedMessage")

For Each objItem in colItems
    Wscript.Echo "Action deleted NDR: " & objItem.ActionDeletedNDR
    Wscript.Echo "Action deleted no NDR: " & _
        objItem.ActionDeletedNoNDR
    Wscript.Echo "Action freeze: " & objItem.ActionFreeze
    Wscript.Echo "Action thaw: " & objItem.ActionThaw
    Wscript.Echo "Expiry: " & objItem.Expiry
    Wscript.Echo "High priority: " & objItem.HighPriority
    Wscript.Echo "Link ID: " & objItem.LinkID
    Wscript.Echo "Link name: " & objItem.LinkName
    Wscript.Echo "Low priority: " & objItem.LowPriority
    Wscript.Echo "Message ID: " & objItem.MessageID
    Wscript.Echo "Normal priority: " & objItem.NormalPriority
    Wscript.Echo "Protocol ID: " & objItem.ProtocolName
    Wscript.Echo "Queue ID: " & objItem.QueueID
    Wscript.Echo "Queue name: " & objItem.QueueName
    Wscript.Echo "Received: " & objItem.Received
    Wscript.Echo "Recipient count: " & objItem.RecipientCount
    Wscript.Echo "Recipients: " & objItem.Recipients
    Wscript.Echo "Sender: " & objItem.Sender
    Wscript.Echo "Size: " & objItem.Size
    Wscript.Echo "State flags: " & objItem.StateFlags
    Wscript.Echo "State frozen: " & objItem.StateFrozen
    Wscript.Echo "State retry: " & objItem.StateRetry
    Wscript.Echo "Subject: " & objItem.Subject
    Wscript.Echo "Submission: " & objItem.Submission
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "Virtual machine: " & objItem.VirtualMachine
    Wscript.Echo "Virtual server name: " & _
        objItem.VirtualServerName
    Wscript.Echo
Next

' List Exchange Queue Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery("Select * from ExchangeQueue")

For Each objItem in colItems
    Wscript.Echo "Can enumerate all: " & objItem.CanEnumAll
    Wscript.Echo "Can enumerate failures: " & objItem.CanEnumFailed
    Wscript.Echo "Van enumerate first N messages: " & _
        objItem.CanEnumFirstNMessages
    Wscript.Echo "Can enumerate frozen messages: " & _
        objItem.CanEnumFrozen
    Wscript.Echo "Can enumerate messages not meeting the criteria: " & _
        objItem.CanEnumInvertSense
    Wscript.Echo "Can enumerate messages larger than X: " & _
        objItem.CanEnumLargerThan
    Wscript.Echo "Can enumerate largest N messages: " & _
        objItem.CanEnumNLargestMessages
    Wscript.Echo "Can enumerate oldest N messages: " & _
        objItem.CanEnumNOldestMessages
    Wscript.Echo "Can enumerate messages older than X: " & _
        objItem.CanEnumOlderThan
    Wscript.Echo "Can enumerate recipients: " & _
        objItem.CanEnumRecipient
    Wscript.Echo "Can enumerate senders: " & objItem.CanEnumSender
    Wscript.Echo "Global stop: " & objItem.GlobalStop
    Wscript.Echo "Increasing time: " & objItem.IncreasingTime
    Wscript.Echo "Link name: " & objItem.LinkName
    Wscript.Echo "Message enumeration flags supported: " & _
        objItem.MsgEnumFlagsSupported
    Wscript.Echo "Number of messages: " & objItem.NumberOfMessages
    Wscript.Echo "Protocol name: " & objItem.ProtocolName
    Wscript.Echo "Queue name: " & objItem.QueueName
    Wscript.Echo "Size of queue: " & objItem.SizeOfQueue
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo "Virtual machine: " & objItem.VirtualMachine
    Wscript.Echo "Virtual server name: " & objItem.VirtualServerName
    Wscript.Echo
Next

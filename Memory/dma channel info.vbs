' List DMA Channel Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_DMAChannel")

For Each objItem in colItems
    Wscript.Echo "Address Size: " & objItem.AddressSize
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Byte Mode: " & objItem.ByteMode
    Wscript.Echo "Channel Timing: " & objItem.ChannelTiming
    Wscript.Echo "DMA Channel: " & objItem.DMAChannel
    Wscript.Echo "Maximum Transfer Size: " & objItem.MaxTransferSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Type C Timing: " & objItem.TypeCTiming
    Wscript.Echo "Word Mode: " & objItem.WordMode
    Wscript.Echo
Next

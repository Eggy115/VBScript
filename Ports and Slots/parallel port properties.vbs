' List Parallel Port Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ParallelPort")

For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    For Each strCapability in objItem.Capabilities
        Wscript.Echo "Capability: " & strCapability
    Next
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OS Auto Discovered: " & objItem.OSAutoDiscovered
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Protocol Supported: " & objItem.ProtocolSupported
Next

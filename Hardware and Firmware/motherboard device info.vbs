' List Motherboard Device Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_MotherboardDevice")

For Each objItem in colItems
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Primary Bus Type: " & objItem.PrimaryBusType
    Wscript.Echo "Secondary Bus Type: " & objItem.SecondaryBusType
    Wscript.Echo
Next


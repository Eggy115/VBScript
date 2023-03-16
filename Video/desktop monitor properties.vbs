' List Desktop Monitor Properties


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor")

For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Bandwidth: " & objItem.Bandwidth
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Display Type: " & objItem.DisplayType
    Wscript.Echo "Is Locked: " & objItem.IsLocked
    Wscript.Echo "Monitor Manufacturer: " & objItem.MonitorManufacturer
    Wscript.Echo "Monitor Type: " & objItem.MonitorType
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Pixels Per X Logical Inch: " & objItem.PixelsPerXLogicalInch
    Wscript.Echo "Pixels Per Y Logical Inch: " & objItem.PixelsPerYLogicalInch
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Screen Height: " & objItem.ScreenHeight
    Wscript.Echo "Screen Width: " & objItem.ScreenWidth
Next

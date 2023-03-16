' List Sound Card Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_SoundDevice")

For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "DMA Buffer Size: " & objItem.DMABufferSize
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "MPU 401 Address: " & objItem.MPU401Address
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Product Name: " & objItem.ProductName
    Wscript.Echo "Status Information: " & objItem.StatusInfo
Next


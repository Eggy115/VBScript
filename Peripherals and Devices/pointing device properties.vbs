' List Pointing Device Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PointingDevice")

For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Device Interface: " & objItem.DeviceInterface
    Wscript.Echo "Double Speed Threshold: " & objItem.DoubleSpeedThreshold
    Wscript.Echo "Handedness: " & objItem.Handedness
    Wscript.Echo "Hardware Type: " & objItem.HardwareType
    Wscript.Echo "INF File Name: " & objItem.InfFileName
    Wscript.Echo "INF Section: " & objItem.InfSection
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Number Of Buttons: " & objItem.NumberOfButtons
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Pointing Type: " & objItem.PointingType
    Wscript.Echo "Quad Speed Threshold: " & objItem.QuadSpeedThreshold
    Wscript.Echo "Resolution: " & objItem.Resolution
    Wscript.Echo "Sample Rate: " & objItem.SampleRate
    Wscript.Echo "Synch: " & objItem.Synch
Next

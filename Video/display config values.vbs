' List Current Display Configuration Values


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_DisplayConfiguration")

For Each objItem in colItems
    Wscript.Echo "Bits Per Pel: " & objItem.BitsPerPel
    Wscript.Echo "Device Name: " & objItem.DeviceName
    Wscript.Echo "Display Flags: " & objItem.DisplayFlags
    Wscript.Echo "Display Frequency: " & objItem.DisplayFrequency
    Wscript.Echo "Driver Version: " & objItem.DriverVersion
    Wscript.Echo "Log Pixels: " & objItem.LogPixels
    Wscript.Echo "Pels Height: " & objItem.PelsHeight
    Wscript.Echo "Pels Width: " & objItem.PelsWidth
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Specification Version: " & objItem.SpecificationVersion
    Wscript.Echo
Next

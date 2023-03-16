' List Keyboard Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Keyboard")
For Each objItem in colItems
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Is Locked: " & objItem.IsLocked
    Wscript.Echo "Layout: " & objItem.Layout
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Number of Function Keys: " & objItem.NumberOfFunctionKeys
    Wscript.Echo "Password: " & objItem.Password
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
Next

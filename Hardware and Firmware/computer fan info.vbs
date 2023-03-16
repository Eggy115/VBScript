



' Retrieving Computer Fan Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Fan")

For Each objItem in colItems
    Wscript.Echo "Active Cooling: " & objItem.ActiveCooling
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Status Information: " & objItem.StatusInfo
    Wscript.Echo
Next

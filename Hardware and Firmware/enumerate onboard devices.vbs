
' Enumerating Onboard Devices


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_OnBoardDevice")

For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device Type: " & objItem.DeviceType
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo
Next

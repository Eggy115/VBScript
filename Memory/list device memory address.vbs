
' List Device Memory Addresses


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_DeviceMemoryAddress")

For Each objItem in colItems
    Wscript.Echo "Ending Address: " & objItem.EndingAddress
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Starting Address: " & objItem.StartingAddress
    Wscript.Echo
Next

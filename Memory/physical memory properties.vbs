' List Physical Memory Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_PhysicalMemoryArray")

For Each objItem in colItems
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Maximum Capacity: " & objItem.MaxCapacity
    Wscript.Echo "Memory Devices: " & objItem.MemoryDevices
    Wscript.Echo "Memory Error Correction: " & objItem.MemoryErrorCorrection
Next

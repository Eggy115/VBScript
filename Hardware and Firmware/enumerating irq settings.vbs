
' Enumerating IRQ Settings


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_IRQResource")

For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Hardware: " & objItem.Hardware
    Wscript.Echo "IRQ Number: " & objItem.IRQNumber
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Trigger Level: " & objItem.TriggerLevel
    Wscript.Echo "Trigger Type: " & objItem.TriggerType
    Wscript.Echo
Next

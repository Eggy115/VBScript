
' List Battery Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Battery")

For Each objItem in colItems
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Battery Status: " & objItem.BatteryStatus
    Wscript.Echo "Chemistry: " & objItem.Chemistry
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Design Voltage: " & objItem.DesignVoltage
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Estimated Run Time: " & objItem.EstimatedRunTime
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Power Management Capabilities: "
    For Each objElement In objItem.PowerManagementCapabilities
        WScript.Echo vbTab & objElement
    Next
    Wscript.Echo "Power Management Supported: " & _
        objItem.PowerManagementSupported
    Wscript.Echo
Next

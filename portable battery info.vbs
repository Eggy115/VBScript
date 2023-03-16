
' List Portable Battery Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PortableBattery")

For Each objItem in colItems
    Wscript.Echo "Capacity Multiplier: " & objItem.CapacityMultiplier
    Wscript.Echo "Chemistry: " & objItem.Chemistry
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Design Capacity: " & objItem.DesignCapacity
    Wscript.Echo "Design Voltage: " & objItem.DesignVoltage
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Location: " & objItem.Location
    dtmWMIDate = objItem.ManufactureDate
    strReturn = WMIDateStringToDate(dtmWMIDate)
    Wscript.Echo "Manufacture Date: " & strReturn
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Maximum Battery Error: " & objItem.MaxBatteryError
    Wscript.Echo "Smart Battery Version: " & objItem.SmartBatteryVersion
    Wscript.Echo
Next
 
Function WMIDateStringToDate(dtmWMIDate)
    If Not IsNull(dtmWMIDate) Then
        WMIDateStringToDate = CDate(Mid(dtmWMIDate, 5, 2) & "/" & _
            Mid(dtmWMIDate, 7, 2) & "/" & Left(dtmWMIDate, 4) _
                & " " & Mid (dtmWMIDate, 9, 2) & ":" & _
                    Mid(dtmWMIDate, 11, 2) & ":" & Mid(dtmWMIDate,13, 2))
    End If
End Function

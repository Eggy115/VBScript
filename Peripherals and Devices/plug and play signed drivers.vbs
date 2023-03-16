
' List Plug and Play Signed Drivers


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PnPSignedDriver")

For Each objItem in colItems
    Wscript.Echo "Class Guid: " & objItem.ClassGuid
    Wscript.Echo "Compatability ID: " & objItem.CompatID
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device Class: " & objItem.DeviceClass
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Device Name: " & objItem.DeviceName
    dtmWMIDate = objItem.DriverDate
    strReturn = WMIDateStringToDate(dtmWMIDate)
    Wscript.Echo "Driver Date: " & strReturn
    Wscript.Echo "Driver Provider Name: " & objItem.DriverProviderName
    Wscript.Echo "Driver Version: " & objItem.DriverVersion
    Wscript.Echo "Hardware ID: " & objItem.HardWareID
    Wscript.Echo "INF Name: " & objItem.InfName
    Wscript.Echo "Is Signed: " & objItem.IsSigned
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "PDO: " & objItem.PDO
    Wscript.Echo "Signer: " & objItem.Signer
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

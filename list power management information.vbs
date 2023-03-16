' List Power Management Information



strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
  If Not IsNull(objItem.PowerManagementCapabilities) Then
    strPowerManagementCapabilities = _
     Join(objItem.PowerManagementCapabilities, ",")
  End If
  WScript.Echo "PowerManagementCapabilities: " & _
   strPowerManagementCapabilities
  WScript.Echo "PowerManagementSupported: " & objItem.PowerManagementSupported
  Select Case objItem.PowerState
    Case 0 strPowerState = "Unknown"
    Case 1 strPowerState = "Full Power"
    Case 2 strPowerState = "Power Save - Low Power Mode"
    Case 3 strPowerState = "Power Save - Standby"
    Case 4 strPowerState = "Power Save - Unknown"
    Case 5 strPowerState = "Power Cycle"
    Case 6 strPowerState = "Power Off"
    Case 7 strPowerState = "Power Save - Warning"
  End Select
  WScript.Echo "PowerState: " & strPowerState
  Select Case objItem.PowerSupplyState
    Case 1 strPowerSupplyState = "Other"
    Case 2 strPowerSupplyState = "Unknown"
    Case 3 strPowerSupplyState = "Safe"
    Case 4 strPowerSupplyState = "Warning"
    Case 5 strPowerSupplyState = "Critical"
    Case 6 strPowerSupplyState = "Non-recoverable"
  End Select
  WScript.Echo "PowerSupplyState: " & strPowerSupplyState
Next

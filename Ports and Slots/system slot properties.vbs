' List System Slot Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_SystemSlot")

For Each objItem in colItems
    For Each strConnectorPinout in objItem.ConnectorPinout
        Wscript.Echo "Connector Pinout: " & strConnectorPinout 
    Next
    Wscript.Echo "Connector Type: " & objItem.ConnectorType
    Wscript.Echo "Current Usage: " & objItem.CurrentUsage
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Height Allowed: " & objItem.HeightAllowed
    Wscript.Echo "Length Allowed: " & objItem.LengthAllowed
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Maximum Data Width: " & objItem.MaxDataWidth
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Number: " & objItem.Number
    Wscript.Echo "PME Signal: " & objItem.PMESignal
    Wscript.Echo "Shared: " & objItem.Shared
    Wscript.Echo "Slot Designation: " & objItem.SlotDesignation
    Wscript.Echo "Supports Hot Plug: " & objItem.SupportsHotPlug
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "Thermal Rating: " & objItem.ThermalRating
    For Each strVccVoltageSupport in objItem.VccMixedVoltageSupport
        Wscript.Echo "VCC Mixed Voltage Support: " & strVccVoltageSupport 
    Next 
    Wscript.Echo "Version: " & objItem.Version
    For Each strVppVoltageSupport in objItem.VppMixedVoltageSupport
        Wscript.Echo "VPP Mixed Voltage Support: " & strVppVoltageSupport 
    Next 
Next

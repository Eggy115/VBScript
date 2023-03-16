' List Port Connector Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PortConnector")

For Each objItem in colItems
    Wscript.Echo "Connector Pinout: " & objItem.ConnectorPinout
    For Each strConnectorType in objItem.ConnectorType
        Wscript.Echo "Connector Type: " & strConnectorType
    Next
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "External Reference Designator: " & _
        objItem.ExternalReferenceDesignator
    Wscript.Echo "Internal Reference Designator: " & _
        objItem.InternalReferenceDesignator
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Port Type: " & objItem.PortType
    Wscript.Echo "Serial Number: " & objItem.SerialNumber
    Wscript.Echo "Tag: " & objItem.Tag
    Wscript.Echo "Version: " & objItem.Version
Next


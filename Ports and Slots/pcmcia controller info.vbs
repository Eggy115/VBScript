' List PCMCIA Controller Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_PCMCIAController")

For Each objItem in colItems
    Wscript.Echo "Configuration Manager Error Code: " & _
        objItem.ConfigManagerErrorCode
    Wscript.Echo "Configuration Manager User Configuration: " & _
        objItem.ConfigManagerUserConfig
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Protocol Supported: " & objItem.ProtocolSupported
    Wscript.Echo
Next

' List Modem Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_POTSModem")

For Each objItem in colItems
    Wscript.Echo "Attached To: " & objItem.AttachedTo
    Wscript.Echo "Blind Off: " & objItem.BlindOff
    Wscript.Echo "Blind On: " & objItem.BlindOn
    Wscript.Echo "Compression Off: " & objItem.CompressionOff
    Wscript.Echo "Compression On: " & objItem.CompressionOn
    Wscript.Echo "Configuration Manager Error Code: " & _
        objItem.ConfigManagerErrorCode
    Wscript.Echo "Configuration Manager User Configuration: " & _
        objItem.ConfigManagerUserConfig
    Wscript.Echo "Configuration Dialog: " & objItem.ConfigurationDialog
    Wscript.Echo "Country Selected: " & objItem.CountrySelected
    Wscript.Echo "DCB: "
    For Each objElement In objItem.DCB
        WScript.Echo vbTab & objElement
    Next
    Wscript.Echo "Default: "
    For Each objElement In objItem.Default
        WScript.Echo vbTab & objElement
    Next
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Device Type: " & objItem.DeviceType
    Wscript.Echo "Driver Date: " & objItem.DriverDate
    Wscript.Echo "Error Control Forced: " & objItem.ErrorControlForced
    Wscript.Echo "Error Control Off: " & objItem.ErrorControlOff
    Wscript.Echo "Error Control On: " & objItem.ErrorControlOn
    Wscript.Echo "Flow Control Hard: " & objItem.FlowControlHard
    Wscript.Echo "Flow Control Off: " & objItem.FlowControlOff
    Wscript.Echo "Flow Control Soft: " & objItem.FlowControlSoft
    Wscript.Echo "Inactivity Scale: " & objItem.InactivityScale
    Wscript.Echo "Inactivity Timeout: " & objItem.InactivityTimeout
    Wscript.Echo "Index: " & objItem.Index
    Wscript.Echo "Maximum Baud Rate To SerialPort: " & _
        objItem.MaxBaudRateToSerialPort
    Wscript.Echo "Model: " & objItem.Model
    Wscript.Echo "Modem INF Path: " & objItem.ModemInfPath
    Wscript.Echo "Modem INF Section: " & objItem.ModemInfSection
    Wscript.Echo "Modulation Bell: " & objItem.ModulationBell
    Wscript.Echo "Modulation CCITT: " & objItem.ModulationCCITT
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Port SubClass: " & objItem.PortSubClass
    Wscript.Echo "Prefix: " & objItem.Prefix
    Wscript.Echo "Properties: "
    For Each objElement In objItem.Properties
        WScript.Echo vbTab & objElement
    Next
    Wscript.Echo "Provider Name: " & objItem.ProviderName
    Wscript.Echo "Pulse: " & objItem.Pulse
    Wscript.Echo "Reset: " & objItem.Reset
    Wscript.Echo "Responses Key Name: " & objItem.ResponsesKeyName
    Wscript.Echo "Speaker Mode Dial: " & objItem.SpeakerModeDial
    Wscript.Echo "Speaker Mode Off: " & objItem.SpeakerModeOff
    Wscript.Echo "Speaker Mode On: " & objItem.SpeakerModeOn
    Wscript.Echo "Speaker Mode Setup: " & objItem.SpeakerModeSetup
    Wscript.Echo "Speaker Volume High: " & objItem.SpeakerVolumeHigh
    Wscript.Echo "Speaker Volume Info: " & objItem.SpeakerVolumeInfo
    Wscript.Echo "Speaker Volume Low: " & objItem.SpeakerVolumeLow
    Wscript.Echo "Speaker Volume Med: " & objItem.SpeakerVolumeMed
    Wscript.Echo "Status Info: " & objItem.StatusInfo
    Wscript.Echo "Terminator: " & objItem.Terminator
    Wscript.Echo "Tone: " & objItem.Tone
    Wscript.Echo
Next



' List Serial Port Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_SerialPort")

For Each objItem in colItems
    Wscript.Echo "Binary: " & objItem.Binary
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Maximum Baud Rate: " & objItem.MaxBaudRate
    Wscript.Echo "Maximum Input Buffer Size: " & objItem.MaximumInputBufferSize
    Wscript.Echo "Maximum Output Buffer Size: " & _
        objItem.MaximumOutputBufferSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "OS Auto Discovered: " & objItem.OSAutoDiscovered
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Provider Type: " & objItem.ProviderType
    Wscript.Echo "Settable Baud Rate: " & objItem.SettableBaudRate
    Wscript.Echo "Settable Data Bits: " & objItem.SettableDataBits
    Wscript.Echo "Settable Flow Control: " & objItem.SettableFlowControl
    Wscript.Echo "Settable Parity: " & objItem.SettableParity
    Wscript.Echo "Settable Parity Check: " & objItem.SettableParityCheck
    Wscript.Echo "Settable RLSD: " & objItem.SettableRLSD
    Wscript.Echo "Settable Stop Bits: " & objItem.SettableStopBits
    Wscript.Echo "Supports 16-Bit Mode: " & objItem.Supports16BitMode
    Wscript.Echo "Supports DTRDSR: " & objItem.SupportsDTRDSR
    Wscript.Echo "Supports Elapsed Timeouts: " & _
        objItem.SupportsElapsedTimeouts
    Wscript.Echo "Supports Int Timeouts: " & objItem.SupportsIntTimeouts
    Wscript.Echo "Supports Parity Check: " & objItem.SupportsParityCheck
    Wscript.Echo "Supports RLSD: " & objItem.SupportsRLSD
    Wscript.Echo "Supports RTSCTS: " & objItem.SupportsRTSCTS
    Wscript.Echo "Supports Special Characters: " & _
        objItem.SupportsSpecialCharacters
    Wscript.Echo "Supports XOn XOff: " & objItem.SupportsXOnXOff
    Wscript.Echo "Supports XOn XOff Setting: " & objItem.SupportsXOnXOffSet
Next

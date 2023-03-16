' List Serial Port Configuration Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_SerialPortConfiguration")

For Each objItem in colItems
    Wscript.Echo "Abort Read Write On Error: " & objItem.AbortReadWriteOnError
    Wscript.Echo "Baud Rate: " & objItem.BaudRate
    Wscript.Echo "Binary Mode Enabled: " & objItem.BinaryModeEnabled
    Wscript.Echo "Bits Per Byte: " & objItem.BitsPerByte
    Wscript.Echo "Continue XMit On XOff: " & objItem.ContinueXMitOnXOff
    Wscript.Echo "CTS Outflow Control: " & objItem.CTSOutflowControl
    Wscript.Echo "Discard NULL Bytes: " & objItem.DiscardNULLBytes
    Wscript.Echo "DSR Outflow Control: " & objItem.DSROutflowControl
    Wscript.Echo "DSR Sensitivity: " & objItem.DSRSensitivity
    Wscript.Echo "DTR Flow Control Type: " & objItem.DTRFlowControlType
    Wscript.Echo "EOF Character: " & objItem.EOFCharacter
    Wscript.Echo "Error Replace Character: " & objItem.ErrorReplaceCharacter
    Wscript.Echo "Error Replacement Enabled: " & _
        objItem.ErrorReplacementEnabled
    Wscript.Echo "Event Character: " & objItem.EventCharacter
    Wscript.Echo "Is Busy: " & objItem.IsBusy
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Parity: " & objItem.Parity
    Wscript.Echo "Parity Check Enabled: " & objItem.ParityCheckEnabled
    Wscript.Echo "RTS Flow Control Type: " & objItem.RTSFlowControlType
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Stop Bits: " & objItem.StopBits
    Wscript.Echo "XOff Character: " & objItem.XOffCharacter
    Wscript.Echo "XOff XMit Threshold: " & objItem.XOffXMitThreshold
    Wscript.Echo "XOn Character: " & objItem.XOnCharacter
    Wscript.Echo "XOn XMit Threshold: " & objItem.XOnXMitThreshold
    Wscript.Echo "XOn XOff InFlow Control: " & objItem.XOnXOffInFlowControl
    Wscript.Echo "XOn XOff OutFlow Control: " & objItem.XOnXOffOutFlowControl
Next

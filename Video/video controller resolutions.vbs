

' List All Possible Video Controller Resolutions


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from CIM_VideoControllerResolution")

For Each objItem in colItems
    Wscript.Echo "Horizontal Resolution: " & objItem.HorizontalResolution
    Wscript.Echo "Number Of Colors: " & objItem.NumberOfColors
    Wscript.Echo "Refresh Rate: " & objItem.RefreshRate
    Wscript.Echo "Scan Mode: " & objItem.ScanMode
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Vertical Resolution: " & objItem.VerticalResolution
    Wscript.Echo
Next

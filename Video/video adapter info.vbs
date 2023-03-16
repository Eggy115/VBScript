
' List Video Adapter Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_DisplayControllerConfiguration")

For Each objItem in colItems
    Wscript.Echo "Bits Per Pixel: " & objItem.BitsPerPixel
    Wscript.Echo "Color Planes: " & objItem.ColorPlanes
    Wscript.Echo "Device Entries in a Color Table: " & _
        objItem.DeviceEntriesInAColorTable
    Wscript.Echo "Device Specific Pens: " & objItem.DeviceSpecificPens
    Wscript.Echo "Horizontal Resolution: " & objItem.HorizontalResolution
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Refresh Rate: " & objItem.RefreshRate
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Vertical Resolution: " & objItem.VerticalResolution
    Wscript.Echo "Video Mode: " & objItem.VideoMode
    Wscript.Echo
Next

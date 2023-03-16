' List Video Controller Properties


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_VideoController")

For Each objItem in colItems
    For Each strCapability in objItem.AcceleratorCapabilities
        Wscript.Echo "Accelerator Capability: " & strCapability
    Next
    Wscript.Echo "Adapter Compatibility: " & objItem.AdapterCompatibility
    Wscript.Echo "Adapter DAC Type: " & objItem.AdapterDACType
    Wscript.Echo "Adapter RAM: " & objItem.AdapterRAM
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Color Table Entries: " & objItem.ColorTableEntries
    Wscript.Echo "Current Bits Per Pixel: " & objItem.CurrentBitsPerPixel
    Wscript.Echo "Current Horizontal Resolution: " & _
        objItem.CurrentHorizontalResolution
    Wscript.Echo "Current Number of Colors: " & objItem.CurrentNumberOfColors
    Wscript.Echo "Current Number of Columns: " & objItem.CurrentNumberOfColumns
    Wscript.Echo "Current Number of Rows: " & objItem.CurrentNumberOfRows
    Wscript.Echo "Current Refresh Rate: " & objItem.CurrentRefreshRate
    Wscript.Echo "Current Scan Mode: " & objItem.CurrentScanMode
    Wscript.Echo "Current Vertical Resolution: " & _
        objItem.CurrentVerticalResolution
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Device Specific Pens: " & objItem.DeviceSpecificPens
    Wscript.Echo "Dither Type: " & objItem.DitherType
    Wscript.Echo "Driver Date: " & objItem.DriverDate
    Wscript.Echo "Driver Version: " & objItem.DriverVersion
    Wscript.Echo "ICM Intent: " & objItem.ICMIntent
    Wscript.Echo "ICM Method: " & objItem.ICMMethod
    Wscript.Echo "INF Filename: " & objItem.InfFilename
    Wscript.Echo "INF Section: " & objItem.InfSection
    Wscript.Echo "Installed Display Drivers: " & _
        objItem.InstalledDisplayDrivers
    Wscript.Echo "Maximum Memory Supported: " & objItem.MaxMemorySupported
    Wscript.Echo "Maximum Number Controlled: " & objItem.MaxNumberControlled
    Wscript.Echo "Maximum Refresh Rate: " & objItem.MaxRefreshRate
    Wscript.Echo "Minimum Refresh Rate: " & objItem.MinRefreshRate
    Wscript.Echo "Monochrome: " & objItem.Monochrome
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Number of Color Planes: " & objItem.NumberOfColorPlanes
    Wscript.Echo "Number of Video Pages: " & objItem.NumberOfVideoPages
    Wscript.Echo "PNP Device ID: " & objItem.PNPDeviceID
    Wscript.Echo "Reserved System Palette Entries: " & _
        objItem.ReservedSystemPaletteEntries
    Wscript.Echo "Specification Version: " & objItem.SpecificationVersion
    Wscript.Echo "System Palette Entries: " & objItem.SystemPaletteEntries
    Wscript.Echo "Video Architecture: " & objItem.VideoArchitecture
    Wscript.Echo "Video Memory Type: " & objItem.VideoMemoryType
    Wscript.Echo "Video Mode: " & objItem.VideoMode
    Wscript.Echo "Video Mode Description: " & objItem.VideoModeDescription
    Wscript.Echo "Video Processor: " & objItem.VideoProcessor
Next


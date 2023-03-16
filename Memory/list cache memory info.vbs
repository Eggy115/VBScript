' List Cache Memory Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_CacheMemory")

For Each objItem in colItems
    Wscript.Echo "Access: " & objItem.Access
    Wscript.Echo "Additional Error Data: "
    For Each objElement In objItem.AdditionalErrorData
        WScript.Echo vbTab & objElement
    Next
    Wscript.Echo "Associativity: " & objItem.Associativity
    Wscript.Echo "Availability: " & objItem.Availability
    Wscript.Echo "Block Size: " & objItem.BlockSize
    Wscript.Echo "Cache Speed: " & objItem.CacheSpeed
    Wscript.Echo "Cache Type: " & objItem.CacheType
    Wscript.Echo "Current SRAM: "
    For Each objElement In objItem.CurrentSRAM
        WScript.Echo vbTab & objElement
    Next
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Device ID: " & objItem.DeviceID
    Wscript.Echo "Error Correct Type: " & objItem.ErrorCorrectType
    Wscript.Echo "Installed Size: " & objItem.InstalledSize
    Wscript.Echo "Level: " & objItem.Level
    Wscript.Echo "Location: " & objItem.Location
    Wscript.Echo "Maximum Cache Size: " & objItem.MaxCacheSize
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Number Of Blocks: " & objItem.NumberOfBlocks
    Wscript.Echo "Status Information: " & objItem.StatusInfo
    Wscript.Echo "Supported SRAM: "
    For Each objElement In objItem.SupportedSRAM
        WScript.Echo vbTab & objElement
    Next
    Wscript.Echo "Write Policy: " & objItem.WritePolicy
    Wscript.Echo
Next


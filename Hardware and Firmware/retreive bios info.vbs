' Retrieving BIOS Information


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colBIOS = objWMIService.ExecQuery _
    ("Select * from Win32_BIOS")

For each objBIOS in colBIOS
    Wscript.Echo "Build Number: " & objBIOS.BuildNumber
    Wscript.Echo "Current Language: " & objBIOS.CurrentLanguage
    Wscript.Echo "Installable Languages: " & objBIOS.InstallableLanguages
    Wscript.Echo "Manufacturer: " & objBIOS.Manufacturer
    Wscript.Echo "Name: " & objBIOS.Name
    Wscript.Echo "Primary BIOS: " & objBIOS.PrimaryBIOS
    Wscript.Echo "Release Date: " & objBIOS.ReleaseDate
    Wscript.Echo "Serial Number: " & objBIOS.SerialNumber
    Wscript.Echo "SMBIOS Version: " & objBIOS.SMBIOSBIOSVersion
    Wscript.Echo "SMBIOS Major Version: " & objBIOS.SMBIOSMajorVersion
    Wscript.Echo "SMBIOS Minor Version: " & objBIOS.SMBIOSMinorVersion
    Wscript.Echo "SMBIOS Present: " & objBIOS.SMBIOSPresent
    Wscript.Echo "Status: " & objBIOS.Status
    Wscript.Echo "Version: " & objBIOS.Version
    For i = 0 to Ubound(objBIOS.BiosCharacteristics)
        Wscript.Echo "BIOS Characteristics: " & _
            objBIOS.BiosCharacteristics(i)
    Next
Next

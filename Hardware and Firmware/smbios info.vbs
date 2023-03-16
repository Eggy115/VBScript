' List SMBIOS Information


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colSMBIOS = objWMIService.ExecQuery _
    ("Select * from Win32_SystemEnclosure")

For Each objSMBIOS in colSMBIOS
    Wscript.Echo "Part Number: " & objSMBIOS.PartNumber
    Wscript.Echo "Serial Number: " & objSMBIOS.SerialNumber
    Wscript.Echo "Asset Tag: " & objSMBIOS.SMBIOSAssetTag
Next

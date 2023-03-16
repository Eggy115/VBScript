' List System Locations



strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colOSItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objOSItem In colOSItems
  WScript.Echo "Boot Device: " & objOSItem.BootDevice
  WScript.Echo "System Device: " & objOSItem.SystemDevice
  WScript.Echo "System Drive: " & objOSItem.SystemDrive
  WScript.Echo "Windows Directory: " & objOSItem.WindowsDirectory
  WScript.Echo "System Directory: " & objOSItem.SystemDirectory
Next

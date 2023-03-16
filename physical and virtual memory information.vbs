
' List Physical and Virtual Memory Information



strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colCSItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objCSItem In colCSItems
  WScript.Echo "Total Physical Memory: " & objCSItem.TotalPhysicalMemory
Next
Set colOSItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objOSItem In colOSItems
  WScript.Echo "Free Physical Memory: " & objOSItem.FreePhysicalMemory
  WScript.Echo "Total Virtual Memory: " & objOSItem.TotalVirtualMemorySize
  WScript.Echo "Free Virtual Memory: " & objOSItem.FreeVirtualMemory
  WScript.Echo "Total Visible Memory Size: " & objOSItem.TotalVisibleMemorySize
Next

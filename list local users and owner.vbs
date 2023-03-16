' List Local Users and Owner



strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colCSItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objCSItem In colCSItems
  WScript.Echo "User Name: " & objCSItem.UserName
  WScript.Echo "Primary Owner Name: " & objCSItem.PrimaryOwnerName
  WScript.Echo "Primary Owner Contact: " & objCSItem.PrimaryOwnerContact
Next
Set colOSItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objOSItem In colOSItems
  WScript.Echo "Registered User: " & objOSItem.RegisteredUser
  WScript.Echo "Number Of Users: " & objOSItem.NumberOfUsers
  WScript.Echo "Number Of Licensed Users: " & objOSItem.NumberOfLicensedUsers
Next

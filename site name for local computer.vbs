' List the Site Name for the Local Computer


Set objADSysInfo = CreateObject("ADSystemInfo")

WScript.Echo "Current site name: " & objADSysInfo.SiteName

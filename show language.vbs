' List Locale and Language Information



strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_OperatingSystem")
For Each objItem In colItems
  WScript.Echo "Country Code: " & objItem.CountryCode
  WScript.Echo "Locale: " & objItem.Locale
  WScript.Echo "OS Language: " & objItem.OSLanguage
  WScript.Echo "Code Set: " & objItem.CodeSet
Next

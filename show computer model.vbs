' List Computer Manufacturer and Model



strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
  WScript.Echo "Name: " & objItem.Name
  WScript.Echo "Manufacturer: " & objItem.Manufacturer
  WScript.Echo "Model: " & objItem.Model
Next

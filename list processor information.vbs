' List Processor Information



strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colCSes = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objCS In colCSes
  WScript.Echo "Computer Name: " & objCS.Name
  WScript.Echo "System Type: " & objCS.SystemType
  WScript.Echo "Number Of Processors: " & objCS.NumberOfProcessors
Next
Set colProcessors = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objProcessor in colProcessors
  WScript.Echo "Manufacturer: " & objProcessor.Manufacturer
  WScript.Echo "Name: " & objProcessor.Name
  WScript.Echo "Description: " & objProcessor.Description
  WScript.Echo "Processor ID: " & objProcessor.ProcessorID
  WScript.Echo "Address Width: " & objProcessor.AddressWidth
  WScript.Echo "Data Width: " & objProcessor.DataWidth
  WScript.Echo "Family: " & objProcessor.Family
  WScript.Echo "Maximum Clock Speed: " & objProcessor.MaxClockSpeed
Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2")
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each objItem In colItems
  WScript.Echo "Computer Name: " & objItem.Name
  WScript.Echo "Name Format: " & objItem.NameFormat
  WScript.Echo "Domain: " & objItem.Domain
  WScript.Echo "Part Of Domain: " & objItem.PartOfDomain 'post-Windows 2000 only
  WScript.Echo "Workgroup: " & objItem.Workgroup 'post-Windows 2000 only
  Select Case objItem.DomainRole
    Case 0 strDomainRole = "Standalone Workstation"
    Case 1 strDomainRole = "Member Workstation"
    Case 2 strDomainRole = "Standalone Server"
    Case 3 strDomainRole = "Member Server"
    Case 4 strDomainRole = "Backup Domain Controller"
    Case 5 strDomainRole = "Primary Domain Controller"
  End Select
  WScript.Echo "Domain Role: " & strDomainRole
  strRoles = Join(objItem.Roles, ",")
  WScript.Echo "Roles: " & strRoles
  WScript.Echo "Network Server Mode Enabled: " & _
   objItem.NetworkServerModeEnabled
Next

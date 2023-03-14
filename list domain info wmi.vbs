' List Domain Information Using WMI


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_NTDomain")

For Each objItem in colItems
    Wscript.Echo "Client Site Name: " & objItem.ClientSiteName
    Wscript.Echo "DC Site Name: " & objItem.DcSiteName
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "DNS Forest Name: " & objItem.DnsForestName
    Wscript.Echo "Domain Controller Address: " & _
        objItem.DomainControllerAddress
    Wscript.Echo "Domain Controller Address Type: " & _
        objItem.DomainControllerAddressType
    Wscript.Echo "Domain Controller Name: " & objItem.DomainControllerName
    Wscript.Echo "Domain GUID: " & objItem.DomainGuid
    Wscript.Echo "Domain Name: " & objItem.DomainName
    Wscript.Echo "DS Directory Service Flag: " & objItem.DSDirectoryServiceFlag
    Wscript.Echo "DS DNS Controller Flag: " & objItem.DSDnsControllerFlag
    Wscript.Echo "DS DNS Domain Flag: " & objItem.DSDnsDomainFlag
    Wscript.Echo "DS DNS Forest Flag: " & objItem.DSDnsForestFlag
    Wscript.Echo "DS Global Catalog Flag: " & objItem.DSGlobalCatalogFlag
    Wscript.Echo "DS Kerberos Distribution Center Flag: " & _
        objItem.DSKerberosDistributionCenterFlag
    Wscript.Echo "DS Primary Domain Controller Flag: " & _
        objItem.DSPrimaryDomainControllerFlag
    Wscript.Echo "DS Time Service Flag: " & objItem.DSTimeServiceFlag
    Wscript.Echo "DS Writable Flag: " & objItem.DSWritableFlag
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Primary Owner Contact: " & objItem.PrimaryOwnerContact
    Wscript.Echo
Next

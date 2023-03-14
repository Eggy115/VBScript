' List Domain Information for Trust Partners


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\MicrosoftActiveDirectory")

Set colDomainInfo = objWMIService.ExecQuery _
    ("Select * from Microsoft_LocalDomainInfo")

For each objDomain in colDomainInfo
    Wscript.Echo "DNS name: " & objDomain.DNSName
    Wscript.Echo "Flat name: " & objDomain.FlatName
    Wscript.Echo "SID: " & objDomain.SID
    Wscript.Echo "Tree name: " & objDomain.TreeName
    Wscript.Echo "Domain controller name: " & objDomain.DCName
Next

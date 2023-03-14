' List Trust Relationships


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\MicrosoftActiveDirectory")

Set colTrustList = objWMIService.ExecQuery _
    ("Select * from Microsoft_DomainTrustStatus")

For each objTrust in colTrustList
    Wscript.Echo "Trusted domain: " & objTrust.TrustedDomain
    Wscript.Echo "Trust direction: " & objTrust.TrustDirection
    Wscript.Echo "Trust type: " & objTrust.TrustType
    Wscript.Echo "Trust attributes: " & objTrust.TrustAttributes
    Wscript.Echo "Trusted domain controller name: " & objTrust.TrustedDCName
    Wscript.Echo "Trust status: " & objTrust.TrustStatus
    Wscript.Echo "Trust is OK: " & objTrust.TrustIsOK
Next



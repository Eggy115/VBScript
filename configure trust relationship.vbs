' Configure Trust Relationship Properties


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\MicrosoftActiveDirectory")

Set colTrustList = objWMIService.ExecQuery _
    ("Select * from Microsoft_TrustProvider")

For Each objTrust in colTrustList
    objTrust.TrustListLifetime = 25
    objTrust.TrustStatusLifetime = 10
    objTrust.TrustCheckLevel = 1
    objTrust.Put_
Next

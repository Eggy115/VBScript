' Enable a Global Catalog Server


strComputer = "atl-dc-01"
 
Const NTDSDSA_OPT_IS_GC = 1
 
Set objRootDSE = GetObject("LDAP://" & strComputer & "/RootDSE")
strDsServiceDN = objRootDSE.Get("dsServiceName")
Set objDsRoot  = GetObject _
    ("LDAP://" & strComputer & "/" & strDsServiceDN)
intOptions = objDsRoot.Get("options")
 
If (intOptions And NTDSDSA_OPT_IS_GC) = FALSE Then
    objDsRoot.Put "options" , intOptions Or NTDSDSA_OPT_IS_GC
    objDsRoot.Setinfo
End If

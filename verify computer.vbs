' Verify that a Computer is a Global Catalog Server


strComputer = "atl-dc-01"
 
Const NTDSDSA_OPT_IS_GC = 1
 
Set objRootDSE = GetObject("LDAP://" & strComputer & "/rootDSE")
strDsServiceDN = objRootDSE.Get("dsServiceName")
Set objDsRoot  = GetObject("LDAP://" & strComputer & "/" & strDsServiceDN)
intOptions = objDsRoot.Get("options")
 
If intOptions And NTDSDSA_OPT_IS_GC Then
    WScript.Echo strComputer & " is a global catalog server."
Else
    Wscript.Echo strComputer & " is not a global catalog server."
End If

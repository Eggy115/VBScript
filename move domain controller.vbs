
' Move a Domain Controller to a New Active Directory Site


strSourceSiteRDN = "cn=Default-First-Site-Name"
strTargetSiteRDN = "cn=Ga-Atl-Sales"
strDcRDN         = "cn=atl-dc-01"
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strDcPath = "LDAP://" & strDcRDN & ",cn=Servers," & strSourceSiteRDN & _
    ",cn=Sites," & strConfigurationNC
 
strTargetSitePath = "LDAP://cn=Servers," & strTargetSiteRDN & _
    ",cn=Sites," & strConfigurationNC
 
Set objTargetSite = GetObject(strTargetSitePath)
objTargetSite.MoveHere strDcPath, strDcRDN

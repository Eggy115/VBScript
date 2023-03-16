' Rename an Active Directory Site


strOldSiteRDN = "cn=Default-First-Site-Name"
strNewSiteRDN = "cn=Ga-Atl-Sales"
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strSitesContainer = "LDAP://cn=Sites," & strConfigurationNC
strOldSitePath = "LDAP://" & strOldSiteRDN & ",cn=Sites," & strConfigurationNC
 
Set objSitesContainer = GetObject(strSitesContainer)
objSitesContainer.MoveHere strOldSitePath, strNewSiteRDN

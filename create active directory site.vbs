' Create an Active Directory Site


strSiteRDN      = "cn=Ga-Atl-Sales"
strSiteLinkRDN  = "cn=DEFAULTIPSITELINK"
strSiteLinkType = "IP"                      
 
Const ADS_PROPERTY_APPEND = 3
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
strSitesContainer = "LDAP://cn=Sites," & strConfigurationNC
 
Set objSitesContainer = GetObject(strSitesContainer)
 
Set objSite = objSitesContainer.Create("site", strSiteRDN)
objSite.SetInfo
 
Set objLicensingSiteSettings = objSite.Create("licensingSiteSettings", _
    "cn=Licensing Site Settings")
objLicensingSiteSettings.SetInfo
 
Set objNtdsSiteSettings = objSite.Create("nTDSSiteSettings", _
     "cn=NTDS Site Settings")
objNtdsSiteSettings.SetInfo
 
Set objServersContainer = objSite.Create("serversContainer", "cn=Servers")
objServersContainer.SetInfo
 
strSiteLinkPath = "LDAP://" & strSiteLinkRDN & ",cn=" & strSiteLinkType & _
    ",cn=Inter-Site Transports,cn=Sites," & strConfigurationNC
 
Set objSiteLink = GetObject(strSiteLinkPath)
objSiteLink.PutEx ADS_PROPERTY_APPEND, "siteList", _
                  Array(objSite.Get("distinguishedName"))
objSiteLink.SetInfo

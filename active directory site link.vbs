' Create an Active Directory Site Link


strSite1Name    = "Ga-Atl-Sales"
strSite2Name    = "Wa-Red-Sales"
strSiteLinkRDN  = "cn=[" & strSite1Name & "][" & strSite2Name & "]"
intCost         = 100
intReplInterval = 60
strDescription  = "[" & strSite1Name & "][" & strSite2Name & "]"
 
Const ADS_PROPERTY_UPDATE = 2
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strSite1DN = "cn=" & strSite1Name & ",cn=Sites," & strConfigurationNC
strSite2DN = "cn=" & strSite2Name & ",cn=Sites," & strConfigurationNC
 
Set objInterSiteTransports = GetObject("LDAP://" & _
    "cn=IP,cn=Inter-Site Transports,cn=Sites," & strConfigurationNC)
 
Set objSiteLink = objInterSiteTransports.Create("siteLink", strSiteLinkRDN)
objSiteLink.Put "cost",         intCost
objSiteLink.Put "replInterval", intReplInterval
objSiteLink.Put "description",  strDescription

 
objSiteLink.PutEx ADS_PROPERTY_UPDATE, "siteList", _
                  Array(strSite1DN, strSite2DN)
objSiteLink.SetInfo

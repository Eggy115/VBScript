' List the Subnets in an Active Directory Site


strSiteRDN = "cn=Ga-Atl-Sales"
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strSitePath = "LDAP://" & strSiteRDN & ",cn=Sites," & strConfigurationNC
 
Set objSite = GetObject(strSitePath)
 
objSite.GetInfoEx Array("siteObjectBL"), 0
arrSiteObjectBL = objSite.GetEx("siteObjectBL")
 
WScript.Echo strSiteRDN & " Subnets" & vbCrLf & _
    String(Len(strSiteRDN) + 8, "-")
 
For Each strSiteObjectBL In arrSiteObjectBL
    WScript.Echo Split(Split(strSiteObjectBL, ",")(0), "=")(1)
Next

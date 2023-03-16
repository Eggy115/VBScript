' Create an Active Directory Subnet


strSubnetRDN     = "cn=192.168.1.0/26"
strSiteObjectRDN = "cn=Ga-Atl-Sales"
strDescription   = "192.168.1.0/255.255.255.192"
strLocation      = "USA/GA/Atlanta"
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strSiteObjectDN = strSiteObjectRDN & ",cn=Sites," & strConfigurationNC
 
strSubnetsContainer = "LDAP://cn=Subnets,cn=Sites," & strConfigurationNC
 
Set objSubnetsContainer = GetObject(strSubnetsContainer)
 
Set objSubnet = objSubnetsContainer.Create("subnet", strSubnetRDN)
objSubnet.Put "siteObject", strSiteObjectDN
objSubnet.Put "description", strDescription

objSubnet.Put "location", strLocation
objSubnet.SetInfo

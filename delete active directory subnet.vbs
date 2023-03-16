
' Delete an Active Directory Subnet


strSubnetCN = "cn=192.168.1.0/26"
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
strSubnetsContainer = "LDAP://cn=Subnets,cn=Sites," & strConfigurationNC
 
Set objSubnetsContainer = GetObject(strSubnetsContainer)
objSubnetsContainer.Delete "subnet", strSubnetCN

' List Servers in an Active Directory Site


strSiteRDN = "cn=Ga-Atl-Sales"
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strServersPath = "LDAP://cn=Servers," & strSiteRDN & ",cn=Sites," & _
    strConfigurationNC
Set objServersContainer = GetObject(strServersPath)
 
For Each objServer In objServersContainer
    WScript.Echo "Name: " & objServer.Name
Next

' List the Subnets in all Active Directory Sites


Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strSubnetsContainer = "LDAP://cn=Subnets,cn=Sites," & strConfigurationNC
 
Set objSubnetsContainer = GetObject(strSubnetsContainer)
 
objSubnetsContainer.Filter = Array("subnet")
 
Set objHash = CreateObject("Scripting.Dictionary")
 
For Each objSubnet In objSubnetsContainer
    objSubnet.GetInfoEx Array("siteObject"), 0
    strSiteObjectDN = objSubnet.Get("siteObject")
    strSiteObjectName = Split(Split(strSiteObjectDN, ",")(0), "=")(1)
 
    If objHash.Exists(strSiteObjectName) Then
        objHash(strSiteObjectName) = objHash(strSiteObjectName) & "," & _
            Split(objSubnet.Name, "=")(1)
    Else
        objHash.Add strSiteObjectName, Split(objSubnet.Name, "=")(1)
    End If
Next
 
For Each strKey In objHash.Keys
    WScript.Echo strKey & "," & objHash(strKey)
Next

' Create a User Account and Add it to a Group and an OU


Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")
Set objOU = objDomain.Create("organizationalUnit", "ou=Management")
objOU.SetInfo
 
Set objOU = GetObject("LDAP://OU=Management,dc=fabrikam,dc=com")
Set objUser = objOU.Create("User", "cn= AckermanPilar")
objUser.Put "sAMAccountName", "AckermanPila"
objUser.SetInfo
 
Set objOU = GetObject("LDAP://OU=Management,dc=fabrikam,dc=com")
Set objGroup = objOU.Create("Group", "cn=atl-users")
objGroup.Put "sAMAccountName", "atl-users"
objGroup.SetInfo
 
objGroup.Add objUser.ADSPath

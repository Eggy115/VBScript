' Create an OU


Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")

Set objOU = objDomain.Create("organizationalUnit", "ou=Management")
objOU.SetInfo

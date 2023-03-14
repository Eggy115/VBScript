
' Create an OU in an Existing OU


Set objOU1 = GetObject("LDAP://ou=OU1,dc=na,dc=fabrikam,dc=com")

Set objOU2 = objOU1.Create("organizationalUnit", "ou=OU2")
objOU2.SetInfo

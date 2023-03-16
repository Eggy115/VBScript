' Create a User Account


Set objOU = GetObject("LDAP://OU=management,dc=fabrikam,dc=com")

Set objUser = objOU.Create("User", "cn=MyerKen")
objUser.Put "sAMAccountName", "myerken"
objUser.SetInfo

' Create a Contact in Active Directory


Set objOU = GetObject("LDAP://OU=management,dc=fabrikam,dc=com")

Set objUser = objOU.Create("contact", "cn=MyerKen")
objUser.SetInfo

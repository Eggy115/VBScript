

' Delete a User Account from Active Directory


Set objOU = GetObject("LDAP://ou=hr,dc=fabrikam,dc=com")

objOU.Delete "user", "cn=MyerKen"

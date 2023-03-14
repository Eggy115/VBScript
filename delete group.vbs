' Delete a Group from Active Directory


Set objOU = GetObject("LDAP://ou=hr,dc=fabrikam,dc=com")

objOU.Delete "group", "cn=atl-users"

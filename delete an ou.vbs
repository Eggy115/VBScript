
' Delete an OU


Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")

objDomain.Delete "organizationalUnit", "ou=hr"

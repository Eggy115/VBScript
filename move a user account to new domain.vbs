' Move a User Account to a New Domain


Set objOU = GetObject("LDAP://ou=management,dc=na,dc=fabrikam,dc=com")

objOU.MoveHere _
    "LDAP://cn=AckermanPilar,OU=management,dc=fabrikam,dc=com", vbNullString

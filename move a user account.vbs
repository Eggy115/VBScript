
' Move a User Account


Set objOU = GetObject("LDAP://ou=sales,dc=na,dc=fabrikam,dc=com")

objOU.MoveHere _
    "LDAP://cn=BarrAdam,OU=hr,dc=na,dc=fabrikam,dc=com", vbNullString

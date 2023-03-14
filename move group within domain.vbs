
' Move a Group Within a Domain


Set objOU = GetObject("LDAP://cn=Users,dc=NA,dc=fabrikam,dc=com")

objOU.MoveHere "LDAP://cn=atl-users,ou=HR,dc=NA,dc=fabrikam,dc=com", _
    vbNullString

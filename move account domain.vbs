' Move a Computer Account to a New Domain


Set objOU = GetObject("LDAP://cn=Computers,dc=NA,dc=fabrikam,dc=com")

objOU.MoveHere "LDAP://cn=Computer01,cn=Users,dc=fabrikam,dc=com", _
    vbNullString

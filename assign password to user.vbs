' Assign a Password to a User


Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=management,dc=fabrikam,dc=com")

objUser.SetPassword "i5A2sj*!"

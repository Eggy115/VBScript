' Unlock a User Account


Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")

objUser.IsAccountLocked = False
objUser.SetInfo

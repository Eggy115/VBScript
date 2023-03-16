
' Set a User Account So It Never Expires


Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")

objUser.AccountExpirationDate = "01/01/1970"
objUser.SetInfo

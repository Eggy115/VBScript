' Require Users to Change Their Password


Set objUser = GetObject _
    ("LDAP://CN=myerken,OU=management,DC=Fabrikam,DC=com")

objUser.Put "pwdLastSet", 0
objUser.SetInfo

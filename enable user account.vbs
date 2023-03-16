' Enable a User Account


Set objUser = GetObject _
  ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")

objUser.AccountDisabled = FALSE
objUser.SetInfo


' Modify the Expiration Date for a User Account


Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")

objUser.AccountExpirationDate = "03/30/2005"
objUser.SetInfo

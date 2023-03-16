' Disable a User Account


Const ADS_UF_ACCOUNTDISABLE = 2
 
Set objUser = GetObject _
("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
intUAC = objUser.Get("userAccountControl")
 
objUser.Put "userAccountControl", intUAC OR ADS_UF_ACCOUNTDISABLE
objUser.SetInfo

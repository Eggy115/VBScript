' Delete User Account Telephone Attributes


Const ADS_PROPERTY_CLEAR = 1 
 
Set objUser = GetObject _
   ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.PutEx ADS_PROPERTY_CLEAR, "info", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "otherPager", 0
objUser.SetInfo

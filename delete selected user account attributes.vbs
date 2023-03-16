
' Delete Selected User Account Attributes


Const ADS_PROPERTY_CLEAR = 1 
 
Set objUser = GetObject _
   ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.PutEx ADS_PROPERTY_CLEAR, "initials", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "otherTelephone", 0
objUser.SetInfo

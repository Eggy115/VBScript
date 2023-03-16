' Clearing User Account Address Attributes


Const ADS_PROPERTY_CLEAR = 1 

Set objUser = GetObject _
   ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.PutEx ADS_PROPERTY_CLEAR, "streetAddress", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "c", 0
objUser.SetInfo

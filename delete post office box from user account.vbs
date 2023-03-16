' Delete a Post Office Box from a User Account


Const ADS_PROPERTY_DELETE = 4 
 
Set objUser = GetObject _
   ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 
 
objUser.PutEx ADS_PROPERTY_DELETE, "postOfficeBox", Array("2224")
objUser.SetInfo

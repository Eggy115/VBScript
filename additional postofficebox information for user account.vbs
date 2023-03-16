' Add Additional postOfficeBox Information for a User Account


Const ADS_PROPERTY_APPEND = 3 
 
Set objUser = GetObject _
   ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 

objUser.PutEx ADS_PROPERTY_APPEND, "postOfficeBox", Array("2225","2226")
objUser.SetInfo

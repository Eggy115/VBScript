' Enable a User to Log on at Any Time


Const ADS_PROPERTY_CLEAR = 1 

Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objUser.PutEx ADS_PROPERTY_CLEAR, "logonHours", 0
objUser.SetInfo

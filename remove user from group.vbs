
' Remove a User from a Group


Const ADS_PROPERTY_DELETE = 4 
 
Set objGroup = GetObject _
   ("LDAP://cn=Sea-Users,cn=Users,dc=NA,dc=fabrikam,dc=com") 
 
objGroup.PutEx ADS_PROPERTY_DELETE, _
    "member",Array("cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
objGroup.SetInfo


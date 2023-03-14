' Remove All the Members of a Group


Const ADS_PROPERTY_CLEAR = 1 
 
Set objGroup = GetObject _
    ("LDAP://cn=Sea-Users,cn=Users,dc=NA,dc=fabrikam,dc=com") 
 
objGroup.PutEx ADS_PROPERTY_CLEAR, "member", 0
objGroup.SetInfo

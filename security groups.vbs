' Add a User to Two Security Groups


Const ADS_PROPERTY_APPEND = 3
 
Set objGroup = GetObject _
    ("LDAP://cn=Atl-Users,cn=Users,dc=NA,dc=fabrikam,dc=com")
objGroup.PutEx ADS_PROPERTY_APPEND, _
    "member", Array("cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
objGroup.SetInfo
 
Set objGroup = GetObject _
    ("LDAP://cn=NA-Employees,cn=Users,dc=NA,dc=fabrikam,dc=com")  
objGroup.PutEx ADS_PROPERTY_APPEND, _
    "member", Array("cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
objGroup.SetInfo

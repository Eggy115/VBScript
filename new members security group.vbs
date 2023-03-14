' Add New Members to a Security Group


Const ADS_PROPERTY_APPEND = 3 
 
Set objGroup = GetObject _
  ("LDAP://cn=Sea-Users,cn=Users,dc=NA,dc=fabrikam,dc=com") 
 
objGroup.PutEx ADS_PROPERTY_APPEND, "member", _
    Array("cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com", _
        "cn=Executives,ou=Management,dc=NA,dc=fabrikam,dc=com", _ 
            "cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objGroup.SetInfo

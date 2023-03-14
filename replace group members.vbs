
' Replace Group Membership with All-New Members


Const ADS_PROPERTY_UPDATE = 2 
 
Set objGroup = GetObject _
    ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com") 
 
objGroup.PutEx ADS_PROPERTY_UPDATE, "member", _
      Array("cn=YoungRob,ou=R&D,dc=NA,dc=fabrikam,dc=com", _
          "cn=ShenAlan,ou=R&D,dc=NA,dc=fabrikam,dc=com")
objGroup.SetInfo

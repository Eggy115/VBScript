' Assign the Primary Group for a User


Const ADS_PROPERTY_APPEND = 3
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
Set objGroup = GetObject _
    ("LDAP://cn=MgmtUniversal,ou=Management,dc=NA,dc=fabrikam,dc=com")
objGroup.GetInfoEx Array("primaryGroupToken"), 0
intPrimaryGroupToken = objGroup.Get("primaryGroupToken")
 
objGroup.PutEx ADS_PROPERTY_APPEND, _
    "member", Array("cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
objGroup.SetInfo
objUser.Put "primaryGroupID", intPrimaryGroupToken
objUser.SetInfo

' Copy a Published Certificate to a User Account


On Error Resume Next

Const ADS_PROPERTY_UPDATE = 2 
 
Set objUserTemplate = _
    GetObject("LDAP://cn=userTemplate,OU=Management,dc=NA,dc=fabrikam,dc=com")
arrUserCertificates = objUserTemplate.GetEx("userCertificate")
 
Set objUser = _
    GetObject("LDAP://cn=MyerKen,OU=Management,dc=NA,dc=fabrikam,dc=com")
objUser.PutEx ADS_PROPERTY_UPDATE, "userCertificate", arrUserCertificates
objUser.SetInfo

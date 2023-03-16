' Delete Published Certificates from a User Account


On Error Resume Next

Const ADS_PROPERTY_DELETE = 4 
 
Set objUserTemplate = _
    GetObject("LDAP://cn=userTemplate,OU=Management,dc=NA,dc=fabrikam,dc=com")
arrUserCertificates = objUserTemplate.GetEx("userCertificate")
 
Set objUser = _
    GetObject("LDAP://cn=MyerKen,OU=Management,dc=NA,dc=fabrikam,dc=com")

objUser.PutEx ADS_PROPERTY_DELETE, "userCertificate", arrUserCertificates
objUser.SetInfo

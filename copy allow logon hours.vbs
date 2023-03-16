
' Copy Allowed Logon Hours from One Account to Another


On Error Resume Next

Set objUserTemplate = _
    GetObject("LDAP://cn=userTemplate,OU=Management,dc=NA,dc=fabrikam,dc=com")
arrLogonHours = objUserTemplate.Get("logonHours")
 
Set objUser = _
    GetObject("LDAP://cn=MyerKen,OU=Management,dc=NA,dc=fabrikam,dc=com")
objUser.Put "logonHours", arrLogonHours
objUser.SetInfo

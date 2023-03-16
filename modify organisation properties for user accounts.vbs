' Modify Organization Properties for a User Account


Set objUser = GetObject _
    ("LDAP://cn=Myerken,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objUser.Put "title", "Manager"
objUser.Put "department", "Executive Management Team"
objUser.Put "company", "Fabrikam"
objUser.Put "manager", _
    "cn=AckermanPilar,OU=Management,dc=NA,dc=fabrikam,dc=com"   
objUser.SetInfo

Set objUser01 = GetObject _
    ("LDAP://cn=LewJudy,OU=Sales,dc=NA,dc=fabrikam,dc=com")
Set objUser02 = GetObject _
    ("LDAP://cn=AckersKim,OU=Sales,dc=NA,dc=fabrikam,dc=com")

objUser01.Put "manager", objUser.Get("distinguishedName")
objUser02.Put "manager", objUser.Get("distinguishedName")   
objUser01.SetInfo
objUser02.SetInfo

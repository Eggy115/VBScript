
' Delete an otherMobile Phone Number


Const ADS_PROPERTY_DELETE = 4
 
Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.PutEx ADS_PROPERTY_DELETE, _
    "otherMobile", Array("(425) 555-3334") 
objUser.SetInfo

' Add an Additional URL to a User Account


Const ADS_PROPERTY_APPEND = 3 
 
Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.PutEx ADS_PROPERTY_APPEND, _
    "url", Array("http://www.fabrikam.com/policy")
objUser.SetInfo

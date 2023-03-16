
' Add a Route to the Dial-In Properties of a User Account


Const ADS_PROPERTY_APPEND = 3 
 
Set objUser = GetObject _
   ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 
objUser.PutEx ADS_PROPERTY_APPEND, _
    "msRASSavedFramedRoute", _
        Array("128.168.0.0/15 0.0.0.0 5") 
objUser.PutEx ADS_PROPERTY_APPEND, _
    "msRADIUSFramedRoute", _
        Array("128.168.0.0/15 0.0.0.0 5")
objUser.SetInfo

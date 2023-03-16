

' Delete All Dial-In Properties for a User Account


Const ADS_PROPERTY_CLEAR = 1
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
  
objUser.PutEx ADS_PROPERTY_CLEAR, "msNPAllowDialin", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "msNPCallingStationID", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "msNPSavedCallingStationID", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSServiceType", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSCallbackNumber", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "msRASSavedCallbackNumber", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSFramedIPAddress", 0
objUser.PutEx ADS_PROPERTY_CLEAR, "msRASSavedFramedIPAddress", 0 
objUser.PutEx ADS_PROPERTY_CLEAR, "msRADIUSFramedRoute", 0  
objUser.PutEx ADS_PROPERTY_CLEAR, "msRASSavedFramedRoute", 0
objUser.SetInfo

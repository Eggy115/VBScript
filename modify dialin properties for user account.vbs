


' Modify Dial-In Properties for a User Account


Const ADS_PROPERTY_UPDATE = 2
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objUser.Put "msNPAllowDialin", TRUE
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "msNPSavedCallingStationID", Array("555-0100", "555-0111")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "msNPCallingStationID", Array("555-0100", "555-0111")
objUser.Put "msRADIUSServiceType", 4
objUser.Put "msRADIUSCallbackNumber", "555-0112" 
objUser.Put "msRASSavedFramedIPAddress", 167903442
objUser.Put "msRADIUSFramedIPAddress", 167903442 'value of 10.2.0.210
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "msRASSavedFramedRoute", _
        Array("10.1.0.0/16 0.0.0.0 1", "192.168.1.0/24 0.0.0.0 3")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "msRADIUSFramedRoute", _
        Array("10.1.0.0/16 0.0.0.0 1", "192.168.1.0/24 0.0.0.0 3")
objUser.SetInfo

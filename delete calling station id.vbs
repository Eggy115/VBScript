' Delete a Calling Station ID from a User Account


Const ADS_PROPERTY_DELETE = 4 
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 

objUser.PutEx ADS_PROPERTY_DELETE, _
    "msNPSavedCallingStationID", Array("555-0111")
objUser.PutEx ADS_PROPERTY_DELETE, _
    "msNPCallingStationID", Array("555-0111")
objUser.SetInfo

' Modify Address Page Information for a User Account


Const ADS_PROPERTY_UPDATE = 2

Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 
 
objUser.Put "streetAddress", "Building 43" & vbCrLf & "One Microsoft Way"
objUser.Put "l", "Redmond"
objUser.Put "st", "Washington"
objUser.Put "postalCode", "98053"
objUser.Put "c", "US"
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "postOfficeBox", Array("2222", "2223", "2224")
objUser.SetInfo

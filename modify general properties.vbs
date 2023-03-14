' Modify the General Properties of an OU


Const ADS_PROPERTY_UPDATE = 2

Set objContainer = GetObject _
    ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
objContainer.Put "street", "Building 43" & vbCrLf & "One Microsoft Way"
objContainer.Put "l", "Redmond"
objContainer.Put "st", "Washington"
objContainer.Put "postalCode", "98053"
objContainer.Put "c", "US"
objContainer.PutEx ADS_PROPERTY_UPDATE, _
    "description", Array("Sales staff")
objContainer.SetInfo

' Clear the General Properties of an OU


Const ADS_PROPERTY_CLEAR = 1 

Set objContainer = GetObject _
  ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
objContainer.PutEx ADS_PROPERTY_CLEAR, "description", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "street", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "l", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "st", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "postalCode", 0
objContainer.PutEx ADS_PROPERTY_CLEAR, "c", 0
objContainer.SetInfo

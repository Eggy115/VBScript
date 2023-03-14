' Remove an OU Manager


Const ADS_PROPERTY_CLEAR = 1 
 
Set objContainer = GetObject _
  ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")

objContainer.PutEx ADS_PROPERTY_CLEAR, "managedBy", 0
objContainer.SetInfo

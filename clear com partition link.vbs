' Clear the COM+ Partition Link Set of an OU


Const ADS_PROPERTY_CLEAR = 1 

Set objContainer = GetObject _
  ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
objContainer.PutEx ADS_PROPERTY_CLEAR, "msCOM-UserPartitionSetLink", 0
objContainer.SetInfo


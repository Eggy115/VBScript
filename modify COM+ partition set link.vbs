' Modify the COM+ Partition Set Link of an OU


Set objContainer = GetObject _
    ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
objContainer.Put "msCOM-UserPartitionSetLink", _
    "cn=PartitionSet1,cn=ComPartitionSets,cn=System,dc=NA,dc=fabrikam,dc=com"
objContainer.SetInfo

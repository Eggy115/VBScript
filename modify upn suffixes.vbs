' Modify the UPN Suffixes Defined in the Forest


Const ADS_PROPERTY_APPEND = 3 

Set objPartitions = GetObject _
    ("LDAP://cn=Partitions,cn=Configuration,dc=fabrikam,dc=com")
 
objPartitions.PutEx ADS_PROPERTY_APPEND, _
    "upnSuffixes", Array("sa.fabrikam.com","corp.fabrikam.com")
objPartitions.SetInfo

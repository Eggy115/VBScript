' List COM+ Partition Information for a Domain


Set objCOMPartitionSets = GetObject _
    ("LDAP://cn=ComPartitionSets,cn=System,dc=NA,dc=fabrikam,dc=com")
 
For Each objPartitionSet in objCOMPartitionSets
    WScript.Echo "Name: " & objPartitionSet.Name
Next

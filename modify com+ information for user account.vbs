' Modify COM+ Information for a User Account


Set objUser = GetObject _
  ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
objUser.Put "msCOM-UserPartitionSetLink", _
  "cn=PartitionSet1,cn=ComPartitionSets,cn=System,dc=NA,dc=fabrikam,dc=com"
objUser.SetInfo

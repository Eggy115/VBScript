
' Assign a New Manager to an OU


Set objContainer = GetObject _
  ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
objContainer.Put "managedBy", "cn=AkersKim,ou=Sales,dc=NA,dc=fabrikam,dc=com"
objContainer.SetInfo

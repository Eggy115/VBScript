' Remove the Manager of a Group


Const ADS_PROPERTY_CLEAR = 1 
 
Set objGroup = GetObject _
   ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com")

objGroup.PutEx ADS_PROPERTY_CLEAR, "managedBy", 0
objGroup.SetInfo


' Delete Selected Attributes from a User Account


Const ADS_PROPERTY_DELETE = 4
 
Set objUser = GetObject _
   ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.PutEx ADS_PROPERTY_DELETE, _
   "otherTelephone", Array("(425) 555-1213") 
objUser.PutEx ADS_PROPERTY_DELETE, "initials", Array("E.")
objUser.SetInfo

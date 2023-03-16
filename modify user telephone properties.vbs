' Modify User Telephone Properties


Const ADS_PROPERTY_UPDATE = 2 

Set objUser = GetObject _
   ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.Put "homePhone", "(425) 555-1111"
objUser.Put "pager", "(425) 555-2222"
objUser.Put "mobile", "(425) 555-3333"
objUser.Put "facsimileTelephoneNumber", "(425) 555-4444"   
objUser.Put "ipPhone", "5555"
objUser.Put "info", "Please do not call this user account" & _
  " at home unless there is a work-related emergency. Call" & _
  " this user's mobile phone before calling the pager number"
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "otherHomePhone", Array("(425) 555-1112")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "otherPager", Array("(425) 555-2223")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "otherMobile", Array("(425) 555-3334", "(425) 555-3335")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "otherFacsimileTelephoneNumber", Array("(425) 555-4445")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "otherIpPhone", Array("6666")
objUser.SetInfo

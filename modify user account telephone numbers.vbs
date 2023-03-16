
' Modify User Account Telephone Numbers


Const ADS_PROPERTY_UPDATE = 2 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 
 
objUser.Put "homePhone", "(425) 555-0100"
objUser.Put "pager", "(425) 555-0101"
objUser.Put "mobile", "(425) 555-0102"
objUser.Put "facsimileTelephoneNumber", "(425) 555-0103"   
objUser.Put "ipPhone", "5555"
objUser.Put "info", "Please do not call this user account" & _
    " at home unless there is a work-related emergency. Call" & _
    " this user's mobile phone before calling the pager number."
objUser.PutEx ADS_PROPERTY_UPDATE, "otherHomePhone", Array("(425) 555-0110")
objUser.PutEx ADS_PROPERTY_UPDATE, "otherPager", Array("(425) 555-0111")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "otherMobile", Array("(425) 555-0112", "(425) 555-0113")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "otherFacsimileTelephoneNumber", Array("(425) 555-0114")
objUser.PutEx ADS_PROPERTY_UPDATE, "otherIpPhone", Array("5556")
objUser.SetInfo

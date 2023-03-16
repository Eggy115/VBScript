' Modify General User Account Attributes


Const ADS_PROPERTY_UPDATE = 2 
Set objUser = GetObject _
   ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com") 
 
objUser.Put "givenName", "Ken"
objUser.Put "initials", "E."
objUser.Put "sn", "Myer"
objUser.Put "displayName", "Myer, Ken"
objUser.Put "physicalDeliveryOfficeName", "Room 4358" 
objUser.Put "telephoneNumber", "(425) 555-1211"
objUser.Put "mail", "myerken@fabrikam.com"
objUser.Put "wWWHomePage", "http://www.fabrikam.com"  
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "description", Array("Management staff")
objUser.PutEx ADS_PROPERTY_UPDATE, _
    "otherTelephone", Array("(800) 555-1212", "(425) 555-1213")  
objUser.PutEx ADS_PROPERTY_UPDATE, _
     "url", Array("http://www.fabrikam.com/management")
objUser.SetInfo

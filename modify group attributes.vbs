
' Modify Group Attributes


Const ADS_PROPERTY_UPDATE = 2 

Set objGroup = GetObject _
   ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com") 
 
objGroup.Put "sAMAccountName", "Scientist01"
objGroup.Put "mail", "YoungRob@fabrikam.com"
objGroup.Put "info", "Use this group for official communications " & _
  "with scientists who are contracted to work with Contoso.com."
objGroup.PutEx ADS_PROPERTY_UPDATE, _
    "description", Array("Scientist Mailing List")
objGroup.SetInfo

' Add 1000 Sample Users to a Security Group


Const ADS_PROPERTY_APPEND = 3 

Set objRootDSE = GetObject("LDAP://rootDSE")
Set objContainer = GetObject("LDAP://cn=Users," & _
    objRootDSE.Get("defaultNamingContext"))
Set objGroup = objContainer.Create("Group", "cn=Group1")
objGroup.Put "sAMAccountName","Group1"
objGroup.SetInfo 

For i = 1 To 1000
    strDN = ",cn=Users," & objRootDSE.defaultNamingContext
    objGroup.PutEx ADS_PROPERTY_APPEND, "member", _
        Array("cn=UserNo" & i & strDN)
    objGroup.SetInfo
Next
WScript.Echo "Group1 created and 1000 Users added to the group."

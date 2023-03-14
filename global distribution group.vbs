
' Create a Global Distribution Group


Const ADS_GROUP_TYPE_GLOBAL_GROUP = &h2

Set objOU = GetObject("LDAP://ou=R&D,dc=NA,dc=fabrikam,dc=com")
Set objGroup = objOU.Create("Group", "cn=Scientists")

objGroup.Put "sAMAccountName", "scientists"
objGroup.Put "groupType", ADS_GROUP_TYPE_GLOBAL_GROUP
objGroup.SetInfo

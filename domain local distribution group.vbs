' Create a Domain Local Distribution Group


Const ADS_GROUP_TYPE_LOCAL_GROUP = &h4

Set objOU = GetObject("LDAP://ou=HR,dc=NA,dc=fabrikam,dc=com")
Set objGroup = objOU.Create("Group", "cn=Vendors")

objGroup.Put "sAMAccountName", "vendors"
objGroup.Put "groupType", ADS_GROUP_TYPE_LOCAL_GROUP
objGroup.SetInfo

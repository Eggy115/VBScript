' Create a Universal Distribution Group


Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8

Set objOU = GetObject("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
Set objGroup = objOU.Create("Group", "cn=Customers")

objGroup.Put "sAMAccountName", "customers"
objGroup.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP
objGroup.SetInfo

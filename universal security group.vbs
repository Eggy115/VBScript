' Create a Universal Security Group


Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000

Set objOU = GetObject("LDAP://cn=Users,dc=NA,dc=fabrikam,dc=com")
Set objGroup = objOU.Create("Group", "cn=All-Employees")

objGroup.Put "sAMAccountName", "AllEmployees"
objGroup.Put "groupType", ADS_GROUP_TYPE_UNIVERSAL_GROUP Or _
    ADS_GROUP_TYPE_SECURITY_ENABLED
objGroup.SetInfo

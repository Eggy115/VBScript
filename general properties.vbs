
' List the General Properties of a Group



On Error Resume Next

Const ADS_GROUP_TYPE_GLOBAL_GROUP = &h2
Const ADS_GROUP_TYPE_LOCAL_GROUP = &h4
Const ADS_GROUP_TYPE_UNIVERSAL_GROUP = &h8
Const ADS_GROUP_TYPE_SECURITY_ENABLED = &h80000000
 
Set objGroup = GetObject _
    ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com")

WScript.Echo "Name: " & objGroup.Name
WScript.Echo "SAM Account Name: " & objGroup.SAMAccountName
WScript.Echo "Mail: " & objGroup.Mail
WScript.Echo "Info: " & objGroup.Info

intGroupType = objGroup.GroupType
 
If intGroupType AND ADS_GROUP_TYPE_LOCAL_GROUP Then
    WScript.Echo "Group scope: Domain local"
ElseIf intGroupType AND ADS_GROUP_TYPE_GLOBAL_GROUP Then
    WScript.Echo "Group scope: Global"
ElseIf intGroupType AND ADS_GROUP_TYPE_UNIVERSAL_GROUP Then
    WScript.Echo "Group scope: Universal"
Else
    WScript.Echo "Group scope: Unknown"
End If
 
If intGroupType AND ADS_GROUP_TYPE_SECURITY_ENABLED Then
    WScript.Echo "Group type: Security group"
Else
    WScript.Echo "Group type: Distribution group"
End If
 
For Each strValue in objGroup.Description
    WScript.Echo "Description: " & strValue
Next

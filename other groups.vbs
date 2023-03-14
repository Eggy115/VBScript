' List Other Groups a Group Belongs To


On Error Resume Next
 
Set objGroup = GetObject _
    ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com")
objGroup.GetInfo
 
arrMembersOf = objGroup.GetEx("memberOf")
 
WScript.Echo "MembersOf:"
For Each strMemberOf in arrMembersOf
    WScript.Echo strMemberOf
Next

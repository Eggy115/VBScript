' List All the Members of a Group


On Error Resume Next
 
Set objGroup = GetObject _
  ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com")
objGroup.GetInfo
 
arrMemberOf = objGroup.GetEx("member")
 
WScript.Echo "Members:"
For Each strMember in arrMemberOf
    WScript.echo strMember
Next

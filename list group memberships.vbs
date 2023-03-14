' List Group Memberships for All the Users in an OU


On Error Resume Next

Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D

Set objOU = GetObject _
    ("LDAP://cn=Users,dc=NA,dc=fabrikam,dc=com")
  
ObjOU.Filter= Array("user")
 
For Each objUser in objOU
    WScript.Echo objUser.cn & " is a member of: " 
    WScript.Echo vbTab & "Primary Group ID: " & _
        objUser.Get("primaryGroupID")
  
    arrMemberOf = objUser.GetEx("memberOf")
  
    If Err.Number <>  E_ADS_PROPERTY_NOT_FOUND Then
        For Each Group in arrMemberOf
            WScript.Echo vbTab & Group
        Next
    Else
        WScript.Echo vbTab & "memberOf attribute is not set"
        Err.Clear
    End If
    Wscript.Echo 
Next

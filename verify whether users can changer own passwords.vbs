' Verify Whether Users Can Change Their Passwords


Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Const CHANGE_PASSWORD_GUID  = _
   "{ab721a53-1e2f-11d0-9819-00aa0040529b}"

Set objUser = GetObject _
  ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
Set objSD = objUser.Get("nTSecurityDescriptor")
Set objDACL = objSD.DiscretionaryAcl

For Each Ace In objDACL
    If ((Ace.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT) And _
        (LCase(Ace.ObjectType) = CHANGE_PASSWORD_GUID)) Then
            blnEnabled = True
    End If
Next

If blnEnabled Then
    WScript.Echo "The user cannot change his or her password."
Else
    WScript.Echo "The user can change his or her password."
End If

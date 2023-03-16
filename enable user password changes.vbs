' Enable Users to Change Their Passwords


Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Const CHANGE_PASSWORD_GUID  = _
    "{ab721a53-1e2f-11d0-9819-00aa0040529b}"
 
Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
Set objSD   = objUser.Get("nTSecurityDescriptor")
Set objDACL = objSD.DiscretionaryAcl
arrTrustees = Array("nt authority\self", "everyone")
 
For Each strTrustee In arrTrustees
    For Each ace In objDACL
        If(LCase(ace.Trustee) = strTrustee) Then
            If((ace.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT) And _
               (LCase(ace.ObjectType) = CHANGE_PASSWORD_GUID)) Then
                   objDACL.RemoveAce ace
            End If
        End If
    Next
Next
 
objUser.Put "nTSecurityDescriptor", objSD
objUser.SetInfo

' Prevent Users From Changing Their Passwords


Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Const ADS_ACEFLAG_OBJECT_TYPE_PRESENT = &H1
Const CHANGE_PASSWORD_GUID = "{ab721a53-1e2f-11d0-9819-00aa0040529b}"
Const ADS_RIGHT_DS_CONTROL_ACCESS = &H100
 
Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
Set objSD = objUser.Get("ntSecurityDescriptor")
Set objDACL = objSD.DiscretionaryAcl
arrTrustees = array("nt authority\self", "EVERYONE")
 
For Each strTrustee in arrTrustees
    Set objACE = CreateObject("AccessControlEntry")
    objACE.Trustee = strTrustee
    objACE.AceFlags = 0
    objACE.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT
    objACE.Flags = ADS_ACEFLAG_OBJECT_TYPE_PRESENT
    objACE.ObjectType = CHANGE_PASSWORD_GUID
    objACE.AccessMask = ADS_RIGHT_DS_CONTROL_ACCESS
    objDACL.AddAce objACE
Next
 
objSD.DiscretionaryAcl = objDACL
objUser.Put "nTSecurityDescriptor", objSD
objUser. SetInfo

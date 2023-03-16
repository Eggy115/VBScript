
' List Password Attributes for a User Account


Const ADS_UF_PASSWORD_EXPIRED = &h800000
Const ADS_ACETYPE_ACCESS_DENIED_OBJECT = &H6
Const CHANGE_PASSWORD_GUID = "{ab721a53-1e2f-11d0-9819-00aa0040529b}"
 
Set objHash = CreateObject("Scripting.Dictionary")
objHash.Add "ADS_UF_PASSWD_NOTREQD", &h00020
objHash.Add "ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED", &h0080
objHash.Add "ADS_UF_DONT_EXPIRE_PASSWD", &h10000
 
Set objUser = GetObject _
    ("LDAP://CN=MyerKen,OU=management,DC=Fabrikam,DC=com")
intUserAccountControl = objUser.Get("userAccountControl")
 
Set objUserNT = GetObject("WinNT://fabrikam/myerken")
intUserFlags = objUserNT.Get("userFlags")
 
If ADS_UF_PASSWORD_EXPIRED And intUserFlags Then
    blnExpiredFlag = True
    Wscript.Echo "ADS_UF_PASSWORD_EXPIRED is enabled"
Else
    Wscript.Echo "ADS_UF_PASSWORD_EXPIRED is disabled"
End If
 
For Each Key In objHash.Keys
    If objHash(Key) And intUserAccountControl Then 
        WScript.Echo Key & " is enabled"
    Else
        WScript.Echo Key & " is disabled"
  End If
Next
 
Set objSD = objUser.Get("nTSecurityDescriptor")
Set objDACL = objSD.DiscretionaryAcl

For Each Ace In objDACL
    If ((Ace.AceType = ADS_ACETYPE_ACCESS_DENIED_OBJECT) And _
        (LCase(Ace.ObjectType) = CHANGE_PASSWORD_GUID)) Then
            blnACEPresent = True
    End If
Next

If blnACEPresent Then
    Wscript.Echo "ADS_UF_PASSWD_CANT_CHANGE is enabled"
Else
    Wscript.Echo "ADS_UF_PASSWD_CANT_CHANGE is disabled"
End If
 
If blnExpiredFlag = True Then 
    Wscript.echo "pwdLastSet is null"
Else 
    Wscript.echo "pwdLastSet is " & objUser.PasswordLastChanged
End If

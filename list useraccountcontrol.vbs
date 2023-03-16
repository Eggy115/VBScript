' List userAccountControl Values for an Active Directory User Account


Set objHash = CreateObject("Scripting.Dictionary")
 
objHash.Add "ADS_UF_SMARTCARD_REQUIRED", &h40000 
objHash.Add "ADS_UF_TRUSTED_FOR_DELEGATION", &h80000 
objHash.Add "ADS_UF_NOT_DELEGATED", &h100000 
objHash.Add "ADS_UF_USE_DES_KEY_ONLY", &h200000 
objHash.Add "ADS_UF_DONT_REQUIRE_PREAUTH", &h400000 
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
intUAC = objUser.Get("userAccountControl")
 
If objUser.IsAccountLocked = True Then
    Wscript.Echo "ADS_UF_LOCKOUT is enabled"
Else
    Wscript.Echo "ADS_UF_LOCKOUT is disabled"
End If
wscript.echo VBCRLF
 
For Each Key In objHash.Keys
    If objHash(Key) And intUAC Then 
        Wscript.Echo Key & " is enabled"
    Else
        Wscript.Echo Key & " is disabled"
  End If
Next

' List Domain Password Property Attributes


Set objHash = CreateObject("Scripting.Dictionary")
 
objHash.Add "DOMAIN_PASSWORD_COMPLEX", &h1
objHash.Add "DOMAIN_PASSWORD_NO_ANON_CHANGE", &h2
objHash.Add "DOMAIN_PASSWORD_NO_CLEAR_CHANGE", &h4
objHash.Add "DOMAIN_LOCKOUT_ADMINS", &h8
objHash.Add "DOMAIN_PASSWORD_STORE_CLEARTEXT", &h16
objHash.Add "DOMAIN_REFUSE_PASSWORD_CHANGE", &h32
 
Set objDomain = GetObject("LDAP://dc=fabrikam,dc=com")
 
intPwdProperties = objDomain.Get("PwdProperties")
WScript.Echo "Password Properties = " & intPwdProperties
 
For Each Key In objHash.Keys
    If objHash(Key) And intPwdProperties Then 
        WScript.Echo Key & " is enabled"
    Else
        WScript.Echo Key & " is disabled"
    End If
Next

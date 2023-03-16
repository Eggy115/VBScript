' Require a User to Logon on Using a Smartcard


Const ADS_UF_SMARTCARD_REQUIRED = &h40000 

Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
intUAC = objUser.Get("userAccountControl")
 
If (intUAC AND ADS_UF_SMARTCARD_REQUIRED) = 0 Then
    objUser.Put "userAccountControl", intUAC XOR ADS_UF_SMARTCARD_REQUIRED
    objUser.SetInfo
End If

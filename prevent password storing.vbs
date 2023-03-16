' Prevent Passwords from Being Stored Using Reversible Encrypted Text


Const ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED = &H80
 
Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
intUAC = objUser.Get("userAccountControl")
 
If intUAC AND _
    ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED Then
        objUser.Put "userAccountControl", intUAC XOR _
            ADS_UF_ENCRYPTED_TEXT_PASSWORD_ALLOWED
        objUser.SetInfo
End If

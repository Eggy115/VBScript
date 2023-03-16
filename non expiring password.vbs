
' Create a Non-Expiring Password


Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
 
Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
intUAC = objUser.Get("userAccountControl")
 
If ADS_UF_DONT_EXPIRE_PASSWD AND intUAC Then
    Wscript.Echo "Already enabled"
Else
    objUser.Put "userAccountControl", intUAC XOR _
        ADS_UF_DONT_EXPIRE_PASSWD
    objUser.SetInfo
    WScript.Echo "Password never expires is now enabled"
End If

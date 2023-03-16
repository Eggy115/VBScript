
' List the Date That a User Account Expires


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")

dtmAccountExpiration = objUser.AccountExpirationDate 
 
If Err.Number = -2147467259 Or dtmAccountExpiration = "1/1/1970" Then
    WScript.Echo "No account expiration date specified"
Else
    WScript.Echo "Account expiration date: " & objUser.AccountExpirationDate
End If


' List When a Password was Last Changed


Set objUser = GetObject _
    ("LDAP://CN=myerken,OU=management,DC=Fabrikam,DC=com")

dtmValue = objUser.PasswordLastChanged
WScript.Echo "Password last changed: " & dtmValue

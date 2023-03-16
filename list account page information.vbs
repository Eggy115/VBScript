' List Account Page Information for a User Account


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=Myerken,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
WScript.Echo "User Principal Name: " & objUser.userPrincipalName
WScript.Echo "SAM Account Name: " & objUser.sAMAccountName
WScript.Echo "User Workstations: " & objUser.userWorkstations

Set objDomain = GetObject("LDAP://dc=NA,dc=fabrikam,dc=com")
WScript.Echo "Domain controller: " & objDomain.dc

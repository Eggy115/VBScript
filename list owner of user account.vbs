' List the Owner of a User Account


Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
Set objNtSecurityDescriptor = objUser.Get("ntSecurityDescriptor")
WScript.Echo "Owner Tab"
WScript.Echo "Current owner of this item: " & objNtSecurityDescriptor.Owner

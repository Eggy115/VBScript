' List COM+ Information for a User Account


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")

WScript.Echo "COM User Partition Set Link: " & _
    objUser.msCOM-UserPartitionSetLink

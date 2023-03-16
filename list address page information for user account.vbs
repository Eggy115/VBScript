' List Address Page Information for a User Account


On Error Resume Next
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
WScript.Echo "Street Address: " & objUser.streetAddress
WScript.Echo "Locality: " & objUser.l
WScript.Echo "State/province: " & objUser.st
WScript.Echo "Postal Code: " & objUser.postalCode
WScript.Echo "Country: " & objUser.c
 
WScript.Echo "Post Office Boxes:"
For Each strValue in objUser.postOfficeBox
    WScript.echo vbTab & vbTab & strValue
Next

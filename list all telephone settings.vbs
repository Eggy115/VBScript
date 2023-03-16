' List All Telephone Settings for a User Account


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
WScript.Echo "Home Phone: " & objUser.homePhone
WScript.Echo "Pager: " & objUser.pager
WScript.Echo "Mobile phone: " & objUser.mobile
WScript.Echo " IP Phone: " & objUser.ipPhone
WScript.Echo "Information: " & objUser.info
WScript.Echo " Fax Number: " & objUser.facsimileTelephoneNumber
 
WScript.Echo "Other Home Phone:"
For Each strValue in objUser.otherHomePhone
    WScript.Echo strValue
Next
 
WScript.Echo "Other Pager:"
For Each strValue in objUser.otherPager
    WScript.Echo strValue
Next
 
WScript.Echo "oOther Mobile Phone:"
For Each strValue in objUser.otherMobile
    WScript.Echo strValue
Next
 
WScript.Echo "Other IP Phone:"
For Each strValue in objUser.otherIpPhone
    WScript.Echo strValue
Next
 
WScript.Echo "Other Fax Number:"
For Each strValue in objUser.otherFacsimileTelephoneNumber
    WScript.Echo strValue
Next

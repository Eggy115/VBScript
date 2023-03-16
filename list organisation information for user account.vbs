
' List Organization Information for a User Account


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")

WScript.Echo "Title: " & objUser.title
WScript.Echo "Department: " & objUser.department
WScript.Echo "Company: " & objUser.company
WScript.Echo "Manager: " & objUser.manager
 
For Each strValue in objUser.directReports
    WScript.Echo "Direct Reports: " & strValue
Next

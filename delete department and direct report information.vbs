' Delete All Department and Direct Report Information from a User Account


On Error Resume Next

Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
Const ADS_PROPERTY_CLEAR = 1 

Set objUser = GetObject _
   ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com") 
objUser.PutEx ADS_PROPERTY_CLEAR, "department", 0
objUser.SetInfo
 
arrDirectReports = objUser.GetEx("directReports")
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Quit
Else
    For Each strValue in arrDirectReports
        Set objUserSource = GetObject("LDAP://" & strValue)
        objUserSource.PutEx ADS_PROPERTY_CLEAR, "manager", 0
        objUserSource.SetInfo
    Next
End If

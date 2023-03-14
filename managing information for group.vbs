' List the Managed By Information for a Group


On Error Resume Next
 
Set objGroup = GetObject _
    ("LDAP://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com")
 
strManagedBy = objGroup.Get("managedBy")
 
If IsEmpty(strManagedBy) = TRUE Then
    WScript.Echo "No user account is assigned to manage " & _
        "this group."
Else
    Set objUser = GetObject("LDAP://" & strManagedBy)

    Call GetUpdateMemberList
 
    WScript.Echo "Office: " & _
        objUser.physicalDeliveryOfficeName  
    WScript.Echo "Street Address: " & objUser.streetAddress
    WScript.Echo "Locality: " & objUser.l
    WScript.Echo "State/Province: " & objUser.st
    WScript.Echo "Country: " & objUser.c
    WScript.Echo "Telephone Number: " & objUser.telephoneNumber
    WScript.Echo "Fax Number: " & _
        objUser.facsimileTelephoneNumber
End If
 
Sub GetUpdateMemberList
    Const ADS_ACETYPE_ACCESS_ALLOWED_OBJECT = &H5 
    Const Member_SchemaIDGuid = "{BF9679C0-0DE6-11D0-A285-00AA003049E2}"
    Const ADS_RIGHT_DS_WRITE_PROP = &H20
    objUser.GetInfoEx Array("canonicalName"),0
    strCanonicalName = objUser.Get("canonicalName")
    strDomain = Mid(strCanonicalName,1,InStr(1,strCanonicalName,".")-1)
    strSAMAccountName = objUser.Get("sAMAccountName")
 
    Set objNtSecurityDescriptor = objGroup.Get("ntSecurityDescriptor")
    Set objDiscretionaryAcl = objNtSecurityDescriptor.DiscretionaryAcl
 
    blnMatch = False
    For Each objAce In objDiscretionaryAcl
        If LCase(objAce.Trustee) = _
            LCase(strDomain & "\" & strSAMAccountName) AND _
            objAce.ObjectType =  Member_SchemaIDGuid AND _
                objAce.AceType = ADS_ACETYPE_ACCESS_ALLOWED_OBJECT AND _
                    objAce.AccessMask And ADS_RIGHT_DS_WRITE_PROP Then
                        blnMatch = True
        End If  
    Next
    If blnMatch Then 
        WScript.Echo "Manager can update the member list"
    Else
        WScript.Echo "Manager cannot update the member list."
    End If
End Sub


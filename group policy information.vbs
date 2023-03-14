' List Group Policy Information for an OU


On Error Resume Next

Set objContainer = GetObject _
    ("LDAP://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
strGpLink = objContainer.Get("gPLink")
intGpOptions = objContainer.Get("gPOptions")
 
If strGpLink <> " " Then
    arrGpLinkItems = Split(strGpLink,"]")
    For i = UBound(arrGPLinkItems) to LBound(arrGpLinkItems) + 1 Step -1
        arrGPLink = Split(arrGpLinkItems(i-1),";")
        strDNGPLink = Mid(arrGPLink(0),9)
        WScript.Echo GetGPOName
        Select Case arrGPLink(1)
        Case 0
            WScript.Echo "No Override is cleared and the GPO is enabled."
        Case 1
            WScript.Echo "No Override is cleared and the GPO is disabled."
        Case 2
            WScript.Echo "No Override is checked and the GPO is enabled."
        Case 3
            WScript.Echo "No Override is checked and the GPO is disabled."
      End Select
    Next
    WScript.Echo VbCrLf
End If
 
If intGpOptions = 1 Then
    WScript.Echo "Block Policy Inheritance is checked."
Else
    WScript.Echo "Block Policy Inheritance is not checked."
End If
    
Function GetGPOName
    Set objConnection = CreateObject("ADODB.Connection")  
    objConnection.Open "Provider=ADsDSOObject;"   
 
    Set objCommand = CreateObject("ADODB.Command")
    objCommand.ActiveConnection = objConnection
 
    objCommand.CommandText = _
        "<LDAP://cn=Policies,cn=System,dc=NA,dc=fabrikam,dc=com>;;" & _
            "distinguishedName,displayName;onelevel"
    Set objRecordSet = objCommand.Execute
 
    Do Until objRecordSet.EOF
        If objRecordSet.Fields("distinguishedName") = strDNGPLink Then
            GetGPOName = objRecordSet.Fields("displayName")
            objConnection.Close
            Exit Function
      End If
      objRecordSet.MoveNext
    Loop
    objConnection.Close
End Function

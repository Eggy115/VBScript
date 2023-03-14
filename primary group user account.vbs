' List the Primary Group for a User Account


On Error Resume Next

Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
intPrimaryGroupID = objUser.Get("primaryGroupID")
 
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
    "<LDAP://dc=NA,dc=fabrikam,dc=com>;(objectCategory=Group);" & _
        "distinguishedName,primaryGroupToken;subtree"  
Set objRecordSet = objCommand.Execute
  
Do Until objRecordset.EOF
    If objRecordset.Fields("primaryGroupToken") = intPrimaryGroupID Then
        WScript.Echo "Primary group:"
        WScript.Echo objRecordset.Fields("distinguishedName") & _
            " (primaryGroupID: " & intPrimaryGroupID & ")"
    End If
    objRecordset.MoveNext
Loop
 
objConnection.Close

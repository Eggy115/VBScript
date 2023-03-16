

' Search for a User Account in Active Directory


strUserName = "kenmyer"
dtStart = TimeValue(Now())
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"
 
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
 
objCommand.CommandText = _
    "<LDAP://dc=fabrikam,dc=com>;(&(objectCategory=User)" & _
         "(samAccountName=" & strUserName & "));samAccountName;subtree"
  
Set objRecordSet = objCommand.Execute
 
If objRecordset.RecordCount = 0 Then
    WScript.Echo "sAMAccountName: " & strUserName & " does not exist."
Else
    WScript.Echo strUserName & " exists."
End If
 
objConnection.Close

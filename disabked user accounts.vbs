' List All the Disabled User Accounts in Active Directory


Const ADS_UF_ACCOUNTDISABLE = 2
 
Set objConnection = CreateObject("ADODB.Connection")
objConnection.Open "Provider=ADsDSOObject;"
Set objCommand = CreateObject("ADODB.Command")
objCommand.ActiveConnection = objConnection
objCommand.CommandText = _
    "<GC://dc=fabrikam,dc=com>;(objectCategory=User)" & _
        ";userAccountControl,distinguishedName;subtree"  
Set objRecordSet = objCommand.Execute
 
intCounter = 0
Do Until objRecordset.EOF
    intUAC=objRecordset.Fields("userAccountControl")
    If intUAC AND ADS_UF_ACCOUNTDISABLE Then
        WScript.echo objRecordset.Fields("distinguishedName") & " is disabled"
        intCounter = intCounter + 1
    End If
    objRecordset.MoveNext
Loop
 
WScript.Echo VbCrLf & "A total of " & intCounter & " accounts are disabled."
 
objConnection.Close


' List All Domain Controllers


Const ADS_SCOPE_SUBTREE = 2

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand = CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"
Set objCOmmand.ActiveConnection = objConnection
 
objCommand.CommandText = _
    "Select distinguishedName from " & _
        "'LDAP://cn=Configuration,DC=fabrikam,DC=com' " _
            & "where objectClass='nTDSDSA'" 
objCommand.Properties("Page Size") = 1000
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 

Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
 
Do Until objRecordSet.EOF
    Wscript.Echo "Computer Name: " & _
        objRecordSet.Fields("distinguishedName").Value
    objRecordSet.MoveNext
Loop

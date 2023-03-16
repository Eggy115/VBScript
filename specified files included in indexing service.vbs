' List Specific Files Included in the Indexing Service


On Error Resume Next

Set objConnection = CreateObject("ADODB.Connection")
objConnection.ConnectionString = "provider=msidxs;"
objConnection.Properties("Data Source") = "Script Catalog"
objConnection.Open
 
Set objCommand = CreateObject("ADODB.Command")
 
strQuery = "Select Filename from Scope()"
 
Set objRecordSet = objConnection.Execute(strQuery)
 
Do While Not objRecordSet.EOF
    Wscript.Echo objRecordSet("Filename")
    objRecordSet.MoveNext
Loop

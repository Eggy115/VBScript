
' Search for Files Using the Indexing Service


On Error Resume Next

Set objConnection = CreateObject("ADODB.Connection")
objConnection.ConnectionString = "provider=msidxs;"
objConnection.Properties("Data Source") = "Script Catalog"
objConnection.Open
 
Set objCommand = CreateObject("ADODB.Command")
 
strQuery = "Select Filename, Size, Contents from Scope() Where " _
    & "Contains('Win32_NetworkAdapterConfiguration')"
 
Set objRecordSet = objConnection.Execute(strQuery)
 
Do While Not objRecordSet.EOF
    Wscript.Echo objRecordSet("Filename"), objRecordSet("Size")
    objRecordSet.MoveNext
Loop

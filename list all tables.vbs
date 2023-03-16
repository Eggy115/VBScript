' List all tables in a Microsoft Access Database

Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "mydatabase.accdb"

strTableName="Table1"

Set objRecordSet = CreateObject("ADODB.Recordset")
Const adSchemaIndexes = 12
Set objRecordSet = objConn.OpenSchema(adSchemaIndexes,Array(Empty, Empty,Empty, Empty, strTableName))

Do Until objRecordset.EOF
    For Each field In objRecordset.Fields
      WScript.Echo field.Name  & ": " & field.Value
    Next
    Wscript.Echo
    objRecordset.MoveNext
Loop

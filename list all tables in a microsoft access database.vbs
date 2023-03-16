
' List all tables in a Microsoft Access Database

Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "mydatabase.accdb"

strTableName="Table1"

Set objRecordSet = CreateObject("ADODB.Recordset")
Const adSchemaColumns = 4
Set objRecordSet = objConn.OpenSchema(adSchemaColumns,Array(Null, Null, strTableName))

Const adArray = 8192
Const adBigInt = 20
Const adBinary = 128
Const adBoolean = 11
Const adBSTR = 8
Const adChapter = 136
Const adChar = 129
Const adCurrency = 6
Const adDate = 7
Const adDBDate = 133
Const adDBTime = 134
Const adDBTimeStamp = 135
Const adDecimal = 14
Const adDouble = 5
Const adEmpty = 0
Const adError = 10
Const adFileTime = 64
Const adGUID = 72
Const adIDispatch = 9
Const adInteger = 3
Const adIUnknown = 13
Const adLongVarBinary = 205
Const adLongVarChar = 201
Const adLongVarWChar = 203
Const adNumeric = 131
Const adPropVariant = 138
Const adSingle = 4
Const adSmallInt = 2
Const adTinyInt = 16
Const adUnsignedBigInt = 21
Const adUnsignedInt = 19
Const adUnsignedSmallInt = 18
Const adUnsignedTinyInt = 17
Const adUserDefined = 132
Const adVarBinary = 204
Const adVarChar = 200
Const adVariant = 12
Const adVarNumeric = 139
Const adVarWChar = 202
Const adWChar = 130

Do Until objRecordset.EOF
    Wscript.Echo "Column name: " & objRecordset.Fields.Item("COLUMN_NAME")
    Wscript.Echo "Data type: " & objRecordset.Fields.Item("DATA_TYPE")  
    Wscript.Echo
    objRecordset.MoveNext
Loop

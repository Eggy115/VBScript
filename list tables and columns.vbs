
' List Tables and Columns in a Microsoft Access Database 

Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "mydatabase.accdb"

Set objRecordSet = CreateObject("ADODB.Recordset")
Const adSchemaTables = 20
Set objRecordSet = objConn.OpenSchema(adSchemaTables)

Do Until objRecordset.EOF
    If objRecordset("TABLE_TYPE")="TABLE" Then
      WScript.Echo objRecordset("TABLE_NAME")
      displayColumns  objRecordset("TABLE_NAME").Value
      WScript.Echo 
    End If
    objRecordset.MoveNext
Loop

objRecordset.Close
objConn.Close

Sub displayColumns(strTableName)

  Set objRecordSet2 = CreateObject("ADODB.Recordset")
  Const adSchemaColumns = 4
  Set objRecordSet2 = objConn.OpenSchema(adSchemaColumns,Array(Null, Null, strTableName))

  Do Until objRecordset2.EOF
    Wscript.Echo "  " & objRecordset2("COLUMN_NAME") & " " & objRecordset2("DATA_TYPE")  
    objRecordset2.MoveNext
  Loop
  objRecordset2.Close
End Sub


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

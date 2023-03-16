' List All Worksheets in an Excel Woorkbook

Set objConnection = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConnection.Open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "data.xlsx; Extended Properties=""Excel 12.0;"""

Set objRecordSet = CreateObject("ADODB.Recordset")
Const adSchemaTables = 20
Set objRecordSet = objConnection.OpenSchema(adSchemaTables)

Do Until objRecordset.EOF
    WScript.Echo objRecordset.Fields.Item("TABLE_NAME") & " - " &  objRecordset.Fields.Item("TABLE_TYPE")
    objRecordset.MoveNext
Loop

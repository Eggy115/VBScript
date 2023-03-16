'Connect to and Query a Microsoft Access Database

Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "mydatabase.accdb"

Set ors = objConn.Execute("select * from table1")

Do While Not(ors.EOF)
  WScript.Echo ors("field1").Value
  ors.MoveNext
Loop

ors.Close

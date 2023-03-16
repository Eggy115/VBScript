
' Select data from an SQLite database


Set cn = CreateObject( "ADODB.Connection" )

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

cn.Open "Driver={SQLite3 ODBC Driver};Database=" & folder & "test.db;"

Set rs= cn.Execute("select * from user")

Do While Not(rs.EOF)
  WScript.Echo rs("email").Value
  rs.MoveNext
Loop

rs.Close
cn.Close

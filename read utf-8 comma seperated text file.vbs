
'Read a UTF-8 Comma Separated text file
'You should create a schema information file (schema.ini) in the same directory

Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & ";Extended Properties=""Text;"";"

Set ors = objConn.Execute("select * from [cities_utf8.txt]")

Do While Not(ors.EOF)
  WScript.Echo ors("city").Value & " " & ors("Latitude").Value & " " & ors("Longitude").Value
  ors.MoveNext
Loop

ors.Close

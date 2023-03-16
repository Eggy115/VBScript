' Extract Data from Microsoft Excel

Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "data.xlsx; Extended Properties=""Excel 12.0;"";"

Set ors = objConn.Execute("select * from [data$]")

Do While Not(ors.EOF)
  WScript.Echo ors("col1").Value & " " & ors("col2").Value
  ors.MoveNext
Loop

ors.Close

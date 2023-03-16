' Add new Rows to a table in a Microsoft Access Database 

Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "mydatabase.accdb"

Set oRS = CreateObject("ADODB.Recordset")
Const adOpenKeyset = 1
Const adLockOptimistic = 3
oRS.Open "table1", objConn, adOpenKeyset, adLockOptimistic

For i=0 to 10
  oRS.AddNew
  oRS("Field1") = "value_" & Int((100 * Rnd) + 1)
  oRS.Update
Next

oRS.Close

objConn.Close

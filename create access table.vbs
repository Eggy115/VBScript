' Create a new Table in a Microsoft Access Database 

Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & "mydatabase.accdb"

Set catalog = CreateObject("ADOX.Catalog")
Set table = CreateObject("ADOX.Table")

Set catalog.ActiveConnection = objConn

table.Name = "Test_Table"
    
Const adInteger = 3
table.Columns.Append "Field1", adInteger

Const adKeyPrimary = 1
table.Keys.Append "PrimaryKey", adKeyPrimary, "Field1"

catalog.Tables.Append table

objConn.Close


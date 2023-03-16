
' Create a new Microsoft Access Database 

connectionString = "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=c:\mynewdatabase.accdb"

Set catalog = CreateObject("ADOX.Catalog")

Set objConn = catalog.Create(connectionString)

objConn.Close() 

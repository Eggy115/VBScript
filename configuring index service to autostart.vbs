' Configuring the Indexing Service to Autostart


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
objAdminIS.EnableCI(True)

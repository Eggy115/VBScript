' Stop the Indexing Service


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
objAdminIS.Stop()

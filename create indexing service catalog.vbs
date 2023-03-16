' Create an Indexing Service Catalog


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
objAdminIS.Stop()

Set objCatalog = objAdminIS.AddCatalog("Script Catalog","c:\scripts")
objAdminIS.Start()

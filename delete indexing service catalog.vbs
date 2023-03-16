
' Delete an Indexing Service Catalog


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
objAdminIS.Stop()
errResult = objAdminIS.RemoveCatalog("Script Catalog", True)
objAdminIS.Start()

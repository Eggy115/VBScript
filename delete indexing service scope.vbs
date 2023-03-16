' Delete an Indexing Service Scope


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
Set objCatalog = objAdminIS.GetCatalogByName("Script Catalog")
objCatalog.RemoveScope("c:\scripts")

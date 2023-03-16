' Create an Indexing Service Scope


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
Set objCatalog = objAdminIS.GetCatalogByName("Script Catalog")
Set objScope = objCatalog.AddScope("c:\scripts\Indexing Server", False)
objScope.Alias = "Script scope"
objScope.Path = "c:\scripts"

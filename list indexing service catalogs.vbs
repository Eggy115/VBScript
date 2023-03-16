' List Indexing Service Catalogs


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
objCatalog = objAdminIS.FindFirstCatalog()
If (objCatalog) Then
    Set objCatAdm = objAdminIS.GetCatalog()
    Wscript.Echo "Catalog location: " & objCatAdm.CatalogLocation
    Wscript.Echo "Catalog name: " & objCatAdm.CatalogName
    If (objAdminIS.IsRunning) Then 
        Wscript.Echo "Is stopped: " & objCatAdm.IsCatalogStopped
        Wscript.Echo "Is paused: " & objCatAdm.IsCatalogPaused
        Wscript.Echo "Is running: " & objCatAdm.IsCatalogRunning
        Wscript.Echo "Delayed filter count: " & objCatAdm.DelayedFilterCount
        Wscript.Echo "Documents to filter: " & objCatAdm.DocumentsToFilter
        Wscript.Echo "Filtered document count: " & _
            objCatAdm.FilteredDocumentCount
        Wscript.Echo "Fresh test count: " & objCatAdm.FreshTestCount
        Wscript.Echo "Index size: " & objCatAdm.IndexSize
        Wscript.Echo "Percent merge complete: " & objCatAdm.PctMergeComplete
        Wscript.Echo "Pending scan count: " & objCatAdm.PendingScanCount
        Wscript.Echo "Persistent index count: " & _
            objCatAdm.PersistentIndexCount
        Wscript.Echo "Query count: " & objCatAdm.QueryCount
        Wscript.Echo "State info: " & objCatAdm.StateInfo
        Wscript.Echo "Total document count: " & objCatAdm.TotalDocumentCount
        Wscript.Echo "Unique key count: " & objCatAdm.UniqueKeyCount
        Wscript.Echo "Word list count: " & objCatAdm.WordListCount
    End If 
End If
 
Do
    objCatalog = objAdminIS.FindNextCatalog()
    If (objCatalog) Then
        Set objCatAdm = objAdminIS.GetCatalog()
        Wscript.Echo "Catalog location: " & objCatAdm.CatalogLocation
        Wscript.Echo "Catalog name: " & objCatAdm.CatalogName
    If (objAdminIS.IsRunning) Then 
        Wscript.Echo "Is stopped: " & objCatAdm.IsCatalogStopped
        Wscript.Echo "Is paused: " & objCatAdm.IsCatalogPaused
        Wscript.Echo "Is running: " & objCatAdm.IsCatalogRunning
        Wscript.Echo "Delayed filter count: " & objCatAdm.DelayedFilterCount
        Wscript.Echo "Documents to filter: " & objCatAdm.DocumentsToFilter
        Wscript.Echo "Filtered document count: " & _
            objCatAdm.FilteredDocumentCount
        Wscript.Echo "Fresh test count: " & objCatAdm.FreshTestCount
        Wscript.Echo "Index size: " & objCatAdm.IndexSize
        Wscript.Echo "Percent merge complete: " & objCatAdm.PctMergeComplete
        Wscript.Echo "Pending scan count: " & objCatAdm.PendingScanCount
        Wscript.Echo "Persistent index count: " & _
            objCatAdm.PersistentIndexCount
        Wscript.Echo "Query count: " & objCatAdm.QueryCount
        Wscript.Echo "State info: " & objCatAdm.StateInfo
        Wscript.Echo "Total document count: " & objCatAdm.TotalDocumentCount
        Wscript.Echo "Unique key count: " & objCatAdm.UniqueKeyCount
        Wscript.Echo "Word list count: " & objCatAdm.WordListCount
        End If 
    Else
        Exit Do
   End If
Loop

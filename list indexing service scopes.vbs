' List Indexing Service Scopes


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
objCatalog = objAdminIS.FindFirstCatalog()
If (objCatalog) Then
    Set objCatAdm = objAdminIS.GetCatalog()
    Set objScopeAdm = objCatAdm.GetScope()
    Wscript.Echo "Alias: " & objScopeAdm.Alias
    Wscript.Echo "Exclude scope: " & objScopeAdm.ExcludeScope
    Wscript.Echo "Logon: " & objScopeAdm.Logon
    Wscript.Echo "Path: " & objScopeAdm.Path
    Wscript.Echo "Virtual scope: " & objScopeAdm.VirtualScope
End If
 
Do
    objCatalog = objAdminIS.FindNextCatalog()
    If (objCatalog) Then
        Set objCatAdm = objAdminIS.GetCatalog()
        Set objScopeAdm = objCatAdm.GetScope()
        Wscript.Echo "Alias: " & objScopeAdm.Alias
        Wscript.Echo "Exclude scope: " & objScopeAdm.ExcludeScope
        Wscript.Echo "Logon: " & objScopeAdm.Logon
        Wscript.Echo "Path: " & objScopeAdm.Path
        Wscript.Echo "Virtual scope: " & objScopeAdm.VirtualScope
    Else
        Exit Do
    End If
Loop

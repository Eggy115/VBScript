' List Exchange Folder Tree Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_FolderTree")

For Each objItem in colItems
    Wscript.Echo "Administrative group: " & _
        objItem.AdministrativeGroup
    Wscript.Echo "Administrative noe: " & _
        objItem.AdministrativeNote
    Wscript.Echo "Associated public stores: " & _
        objItem.AssociatedPublicStores
    Wscript.Echo "Creation time: " & objItem.CreationTime
    Wscript.Echo "GUID: " & objItem.GUID
    Wscript.Echo "Has local public store: " & _
        objItem.HasLocalPublicStore
    Wscript.Echo "Last modification time: " & _
        objItem.LastModificationTime
    Wscript.Echo "MAPI folder tree: " & objItem.MAPIFolderTree
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Root folder URL: " & objItem.RootFolderURL
    Wscript.Echo
Next

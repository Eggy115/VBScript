' List Items in the Internet Favorites Folder


Const FAVORITES = &H6&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(FAVORITES)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

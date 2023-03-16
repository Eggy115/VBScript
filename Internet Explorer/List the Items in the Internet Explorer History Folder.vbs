' List the Items in the Internet Explorer History Folder


Const LOCAL_SETTINGS_HISTORY = &H22&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(LOCAL_SETTINGS_HISTORY)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

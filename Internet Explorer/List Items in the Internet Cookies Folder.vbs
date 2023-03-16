' List Items in the Internet Cookies Folder


Const COOKIES = &H21&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(COOKIES)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

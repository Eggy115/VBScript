

' List Items in the "Internet Explorer" folder

Const INTERNET_EXPLORER = &H1

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(INTERNET_EXPLORER)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
  Wscript.Echo objItem.Name
Next

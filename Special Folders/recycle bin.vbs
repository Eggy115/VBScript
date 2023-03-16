' List Items in the "Recycle Bin" folder

Const RECYCLE_BIN = &Ha

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(RECYCLE_BIN)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
  Wscript.Echo objItem.Name
Next


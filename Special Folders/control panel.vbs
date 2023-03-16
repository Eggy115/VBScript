' List Items in the "Control Panel" folder

Const CONTROL_PANEL = &H3

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(CONTROL_PANEL)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
  Wscript.Echo objItem.Name
Next



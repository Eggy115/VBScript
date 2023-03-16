' List Items in the "My Computer" folder

Const MY_COMPUTER = &H11

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_COMPUTER)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
  Wscript.Echo objItem.Name
Next

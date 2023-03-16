

' List Items in the "My Network Places" folder

Const MY_NETWORK_PLACES = &H12

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_NETWORK_PLACES)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
  Wscript.Echo objItem.Name
Next

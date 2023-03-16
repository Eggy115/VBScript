' List Items in the "Network Connections" folder

Const NETWORK_CONNECTIONS = &H31

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(NETWORK_CONNECTIONS)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
  Wscript.Echo objItem.Name
Next

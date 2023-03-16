


' List Items in the "Printers and Faxes" folder

Const PRINTERS_AND_FAXES = &H4

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(PRINTERS_AND_FAXES)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
  Wscript.Echo objItem.Name
Next


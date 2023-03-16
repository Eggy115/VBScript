' List Items in the Temporary Internet Files Folder


Const TEMPORARY_INTERNET_FILES = &H20&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(TEMPORARY_INTERNET_FILES)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set colItems = objFolder.Items
For Each objItem in colItems
    Wscript.Echo objItem.Name
Next

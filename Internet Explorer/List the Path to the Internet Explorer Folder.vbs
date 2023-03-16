' List the Path to the Internet Explorer Folder


Const INTERNET_EXPLORER = &H1&

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(INTERNET_EXPLORER)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

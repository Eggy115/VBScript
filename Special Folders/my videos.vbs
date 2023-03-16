' List Items in the "My Videos" folder

Const MY_VIDEOS = &He

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_VIDEOS)
Set objFolderItem = objFolder.Self
Wscript.Echo objFolderItem.Path

Set fso = CreateObject("Scripting.Filesystemobject")

Set folder=fso.GetFolder(objFolderItem.Path)

For Each subfolder in folder.SubFolders
  WScript.Echo "[" & subfolder.Name & "]"
Next

For Each file in folder.Files
  WScript.Echo file.Name
Next

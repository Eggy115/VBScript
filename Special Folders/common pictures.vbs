' List Items in the "Common Pictures" folder

Const CSIDL_COMMON_PICTURES = &H36

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(CSIDL_COMMON_PICTURES)
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

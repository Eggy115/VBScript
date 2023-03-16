' List Items in the "Common Start Menu Programs" folder

Const CSIDL_COMMON_PROGRAMS = &H17

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(CSIDL_COMMON_PROGRAMS)
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

' List Items in the "Program Files - Common" folder

Const CSIDL_PROGRAM_FILES_COMMON = &H2b

Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(CSIDL_PROGRAM_FILES_COMMON)
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

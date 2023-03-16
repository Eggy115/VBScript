' List All Special Folders in Windows 

Set objShell = CreateObject("Shell.Application")

For i=0 to 255

  Set objFolder = objShell.Namespace(i)

  On Error Resume next
  Set objFolderItem = objFolder.Self
  On Error Resume Next
  WScript.Echo i & " " &  objFolder.Title & " " & objFolderItem.Path
  Set objFolder=Nothing
  
Next

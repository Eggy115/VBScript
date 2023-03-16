' List Path to the My Pictures Folders


Const MY_PICTURES = &H27&
Set objShell = CreateObject("Shell.Application")
Set objFolder = objShell.Namespace(MY_PICTURES) 
Set objFolderItem = objFolder.Self      
Wscript.Echo objFolderItem.Name & ": " & objFolderItem.Path

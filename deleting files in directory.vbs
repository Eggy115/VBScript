' Deleting Files in a Directory with Names That Match a Regular Expression



strFolderName = "C:\scripts\textfiles\delete\"
strREPattern = "log\d\d_\d\d05\.txt"

Set objFS = CreateObject("Scripting.FileSystemObject")

Set objFolder = objFS.GetFolder(strFolderName)
Set objRE = new RegExp
objRE.Pattern    = strREPattern
objRE.IgnoreCase = True

For Each objFile In objFolder.Files
   If objRE.Test(objFile.Name) Then
      objFS.DeleteFile(strFolderName & objFile.Name)
   End If
Next

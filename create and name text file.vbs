' Create and Auto-Name a Text File


Set objFSO = CreateObject("Scripting.FileSystemObject")

strPath = "C:\FSO"
strFileName = objFSO.GetTempName
strFullName = objFSO.BuildPath(strPath, strFileName)
Set objFile = objFSO.CreateTextFile(strFullName)
objFile.Close
objFSO.DeleteFile(strFullName)

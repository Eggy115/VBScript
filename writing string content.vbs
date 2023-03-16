
' Writing String Content to End of Existing Text File



Const FOR_APPENDING = 8
strFileName = "C:\scripts\testfile.txt"
strContent  = "sample string content"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName,FOR_APPENDING)
objTS.Write strContent

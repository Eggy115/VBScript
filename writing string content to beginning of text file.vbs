
' Writing String Content to Beginning of Existing Text File



Const FOR_READING = 1
Const FOR_WRITING = 2

strFileName = "C:\scripts\test.txt"
strNewContent  = "sample string content"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName,FOR_READING)
strContents = objTS.ReadAll
objTS.Close
Set objTS = objFS.OpenTextFile(strFileName,FOR_WRITING)
objTS.WriteLine strNewContent
objTS.Write strContents

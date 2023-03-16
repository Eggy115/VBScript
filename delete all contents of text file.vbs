' Delete All Contents of a Text File



Const FOR_WRITING = 2
strFileName = "C:\scripts\test.txt"

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)

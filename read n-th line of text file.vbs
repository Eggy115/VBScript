' Read n-th Line of a Text File



Const FOR_READING = 1
strFilePath = "C:\scripts\test.txt"
iLineNumber = 5

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFilePath, FOR_READING)

For i=1 To (iLineNumber-1)
   objTS.SkipLine
Next

WScript.Echo objTS.Readline

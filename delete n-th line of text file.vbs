' Delete n-th Line of a Text File



Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = "C:\scripts\test.txt"
iLineNumber = 3

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)

arrLines = Split(strContents,vbNewLine)
For i=0 To UBound(arrLines)
   If i<> (iLineNumber-1) Then
      objTS.WriteLine arrLines(i)
   End If
Next

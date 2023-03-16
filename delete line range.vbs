' Delete Line Range within a Text File



Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = "C:\scripts\test.txt"
iStartAtLineNumber = 3
iEndAtLineNumber   = 7

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

arrLines = Split(strContents, vbNewLine)
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)

For i=0 To UBound(arrLines)
   If i < (iStartAtLineNumber-1) OR i > (iEndAtLineNumber-1) Then
      objTS.WriteLine arrLines(i)
   End If
Next

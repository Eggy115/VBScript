' Delete First n Lines of a Text File



Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = "C:\scripts\test.txt"
iNumberOfLinesToDelete = 5

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

arrLines = Split(strContents, vbNewLine)
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)

For i=0 To UBound(arrLines)
   If i > (iNumberOfLinesToDelete - 1) Then
      objTS.WriteLine arrLines(i)
   End If
Next

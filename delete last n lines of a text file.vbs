' Delete Last n Lines of a Text File



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
iIndexToDeleteFrom = UBound(arrLines)- iNumberOfLinesToDelete + 1

For i=0 To UBound(arrLines)
   If i < iIndexToDeleteFrom Then
      objTS.WriteLine arrLines(i)
   End If
Next

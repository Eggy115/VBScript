

' Writing String Content Following n-th Line of Existing Text File



Const FOR_READING = 1
Const FOR_WRITING = 2

strFileName = "C:\scripts\test.txt"
strNewContent = "sample string content"
iInsertAfterLineNumber = 3

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName,FOR_READING)
strContents = objTS.ReadAll
objTS.Close

Set objTS = objFS.OpenTextFile(strFileName,FOR_WRITING)
arrLines  = Split(strContents, vbNewLine)

For i=0 To UBound(arrLines)
   If i = iInsertAfterLineNumber Then
      objTS.WriteLine strNewContent
   End If
   objTS.WriteLine arrLines(i)
Next

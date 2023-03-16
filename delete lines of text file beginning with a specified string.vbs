' Delete Lines of a Text File Beginning with a Specified String



Const FOR_READING = 1
Const FOR_WRITING = 2
strFileName = "C:\scripts\test.txt"
strCheckForString = UCase("july")

Set objFS = CreateObject("Scripting.FileSystemObject")
Set objTS = objFS.OpenTextFile(strFileName, FOR_READING)
strContents = objTS.ReadAll
objTS.Close

arrLines = Split(strContents, vbNewLine)
Set objTS = objFS.OpenTextFile(strFileName, FOR_WRITING)

For Each strLine In arrLines
   If Not(Left(UCase(LTrim(strLine)),Len(strCheckForString)) = strCheckForString) Then  
      objTS.WriteLine strLine
   End If
Next

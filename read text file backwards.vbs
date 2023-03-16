' Read a Text File from the Bottom Up


Dim arrFileLines()
i = 0

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objFile = objFSO.OpenTextFile("C:\FSO\ScriptLog.txt", 1)

Do Until objFile.AtEndOfStream
     Redim Preserve arrFileLines(i)
     arrFileLines(i) = objFile.ReadLine
     i = i + 1
Loop

objFile.Close

For l = Ubound(arrFileLines) to LBound(arrFileLines) Step -1
    Wscript.Echo arrFileLines(l)
Next

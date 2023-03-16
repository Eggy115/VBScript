' Read a Text File Character-by-Character


Set objFSO = CreateObject("Scripting.FileSystemObject")

Set objFile = objFSO.OpenTextFile("C:\FSO\New Text Document.txt", 1)
Do Until objFile.AtEndOfStream
    strCharacters = objFile.Read(1)
    Wscript.Echo strCharacters
Loop

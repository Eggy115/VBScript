

' Read an Entire Text File

Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("c:\temp\myfile.txt", ForReading)

contents = objTextFile.ReadAll
objTextFile.Close

WScript.Echo contents

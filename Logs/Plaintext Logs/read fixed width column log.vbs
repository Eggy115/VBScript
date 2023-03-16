' Read a Fixed Width Column Log


Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("C:\Windows\Debug\Netsetup.log", _
    ForReading)

Do While objTextFile.AtEndOfStream <> True
    strLinetoParse = objTextFile.ReadLine
    dtmEventDate = Mid(strLinetoParse, 1, 6)
    dtmEventTime = Mid(strLinetoParse, 7, 9)
    strEventDescription = Mid(strLinetoParse, 16)
    Wscript.Echo "Date: " & dtmEventDate
    Wscript.Echo "Time: " & dtmEventTime
    Wscript.Echo "Description: " & strEventDescription & VbCrLf
Loop
objFSO.Close

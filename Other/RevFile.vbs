REM ===========================================================================
REM Bakup file to reverse
REM =========================================================================== 

const CONSOLE_HIDE=0
const CONSOLE_SHOW=1
const CMD_WAIT=true

Const ForReading = 1, ForWriting = 2, ForAppending = 8
set oShell = WScript.CreateObject("WScript.shell")
set oFs = CreateObject("Scripting.FileSystemObject")
set objArgs = Wscript.Arguments

If objArgs.Count <> 3 then
	WScript.ECHO "Reserve error args"
	WSCript.Quit 1
End If

'Get src disk
srcDrive = Wscript.Arguments(1)
if not oFs.FolderExists(srcDrive) then
	WScript.echo "Can't find drive [" & srcDrive & "]"
	WScript.Quit 1
end if

'Get output folder
outWhere = Wscript.Arguments(2)
if not oFs.FolderExists(outWhere) then
	oFs.CreateFolder(outWhere)
end if

'Get input log file
inFileName = Wscript.Arguments(0)
set inFile = oFs.OpenTextFile(inFileName, ForReading , True)

Do While Not inFile.AtEndOfStream
	sName = srcDrive & Trim(inFile.ReadLine)

	if (oFs.FileExists(sName)) then
		pos1=InStr(sName,"\")
		pos2=InStrRev(sName,"\")
		strOutFilePath = outWhere & Mid(sName,pos1+1,pos2-pos1)
		WScript.ECHO "Src=" & sName & "           " & "Dest=" & strOutFilePath
		
		set oShell = wscript.createObject("WScript.Shell")
		oShell.run "cmd /c MKDIR " & strOutFilePath, CONSOLE_HIDE, CMD_WAIT
		oFs.CopyFile sName, strOutFilePath
	end if
Loop
inFile.Close

WSCript.Quit 0
 
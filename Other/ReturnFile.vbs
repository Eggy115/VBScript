REM ===========================================================================
REM ReturnFile.vbs
REM Bakup file to reverse
REM 2010-01-06
REM ===========================================================================
const CONSOLE_HIDE=0
const CONSOLE_SHOW=1
const CMD_WAIT=true

Const ForReading = 1, ForWriting = 2, ForAppending = 8
set oShell = WScript.CreateObject("WScript.shell")
set oFs = CreateObject("Scripting.FileSystemObject")
set objArgs = Wscript.Arguments

If objArgs.Count <> 3 then
   WScript.ECHO "Reverse error args"
   WSCript.Quit 1
End If

'Get target disk
outDrive = Wscript.Arguments(2)
if not oFs.FolderExists(outDrive) then
  WScript.echo "Can't find drive [" & outDrive & "]"
  WScript.Quit 1
end if

'Get source folder
srcPath= Wscript.Arguments(1)
if not oFs.FolderExists(srcPath) then

end if

'Get src log file
inFileName = Wscript.Arguments(0)
set inFile = oFs.OpenTextFile(inFileName, ForReading , True)

' Do While Not inFile.AtEndOfStream
'   outWhere = Trim(inFile.ReadLine)
'   pos = InStrRev(outWhere, "\")
'   if (pos > 0) then
'     outPath = outDrive & Left(outWhere, pos)
'
'     srcWhere = srcPath & Replace(outWhere, Left(outWhere, pos) , "")
'     if oFs.FileExists(srcWhere) then
'       wscript.echo "srcWhere= " & srcWhere & " , outPath= " & outPath
'       oFs.CopyFile srcWhere, outPath
'     end if
'   end if
' Loop


Do While Not inFile.AtEndOfStream
  outWhere = Trim(inFile.ReadLine)
  strtmp = outDrive & outWhere
  pos2=InStrRev(strtmp,"\")
  outPath = Mid(strtmp, 1, pos2)
  srcWhere = srcPath & outWhere
  if oFs.FileExists(srcWhere) then
		wscript.echo "srcWhere= " & srcWhere & " , outPath= " & outPath
		set oShell = wscript.createObject("WScript.Shell")
		oShell.run "cmd /c MKDIR " & outPath, CONSOLE_HIDE, CMD_WAIT
		oFs.CopyFile srcWhere, outPath
	end if
Loop

inFile.Close
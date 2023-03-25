'dim driveRP, pathPatch, isDebug
set oArgs = wscript.arguments
if oArgs.count <> 1 then
	wscript.echo "DetectDrv error args"
	wscript.quit 1
end if

CMDPath=chr(34) & oArgs(0) & chr(34)

Dim wshell, CMDResult
'On Error Resume Next
CMDResult=99
set wshell = CreateObject("WScript.Shell")
CMDResult = wshell.run (CMDPath, 0, True)
wscript.quit CMDResult
 

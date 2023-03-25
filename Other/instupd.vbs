On Error Resume Next
const secStep = 10
const defaultTimeoutMaxSeconds = 1800
Const ForReading = 1, ForWriting = 2, ForAppending = 8

Dim objArgs, fso, objShell
Dim InstallBom,InstallFile,CVAFolder,CMDFolder, LogFileName, LogFolder, CVAFile, CVAPath, UIA, DelivName, VersionStr, RevisionStr, PassStr, DelivPN, DelivSize, DriveFreeSize, InstallSize, RtnCode, InstStage, FccResult, sLocale
Dim LogFile : LogFile = "c:\system.sav\logs\instupd.log"
Dim CmdLineTxtSummaryLog : CmdLineTxtSummaryLog = "c:\System.sav\logs\ImgEnh_UIA_Deliverable_CmdLine.log"
Set objArgs = WScript.Arguments
Set fso = CreateObject("Scripting.FileSystemObject")
InstallFile = Wscript.Arguments(0)
InstallBom = "C:\System.sav\Util\Install.bom"
if fso.FileExists(InstallFile) <> True then InstallFile = Replace(InstallFile, "C:", "I:", 1, -1, vbTextCompare)
if fso.FileExists(InstallFile) <> True then WScript.Quit(1)

CVAFolder = fso.GetParentFolderName(InstallFile)
CMDFolder = CVAFolder
WSCript.Echo "[" & Right(CVAFolder, 4) & "]"
if StrComp(Right(CVAFolder, 4), "\src", vbTextCompare)=0 then
	CVAFolder = Left(CVAFolder, Len(CVAFolder)-4)
end if
WSCript.Echo "CVAFolder=[" & CVAFolder & "]"
WSCript.Echo "CMDFolder=[" & CMDFolder & "]"


sLocale = "en-us"
sLocale = getLocale()
SetLocale("en-us")
LOG vbCrLf & vbCrLf & "=========================================================================="
LOG "Current Locale: " & sLocale
CVAFile = GetFileExtend(CVAFolder,"CVA")
CVAPath = CVAFolder & "\" & CVAFile
LOG "CVA File: " & CVAFile
LOG "CVA Path: " & CVAPath
DelivName = ReadIniValue(CVAPath,"Software Title","US","")
UIA = ReadIniValue("C:\system.sav\Logs\MYSYSTEM.INI","SUMMARY","UIAErrorNo","")


LogFileName = "[FCC] " & Trim(Left(DelivName, 80))
LOG "LogFileName=" & LogFileName

rem Legal chars but try to refine them: space,(,),!,&
rem Illegal chars for file name: \,/,:,*,?,<,>,|,""
Dim ilglC
for each ilglC in Split(" ,(,),!,&,\,/,:,*,?,<,>,|,""", ",")
	LOG "Replacing from '" & ilglC & "' to '_'"
	LogFileName = Replace(LogFileName, ilglC, "_", 1, -1, vbTextCompare)
Next
LOG "LogFileName=" & LogFileName
LOG "LogFile=" & "C:\System.sav\Logs\BB\" & LogFileName & ".LOG"
LogFile =        "C:\System.sav\Logs\BB\" & LogFileName & ".LOG"

LogFolder = "C:\System.sav\Logs\" & LogFileName
LOG "LogFolder=" & LogFolder
if not fso.FolderExists(LogFolder) then fso.CreateFolder(LogFolder)

LOG "CVAFolder=[" & CVAFolder & "]"
LOG "CMDFolder=[" & CMDFolder & "]"


LOG "Install Folder: " & CVAFolder
VersionStr = Replace(ReadIniValue(CVAPath,"General","Version",""),chr(9),"")
RevisionStr = Replace(ReadIniValue(CVAPath,"General","Revision",""),chr(9),"")
PassStr = Replace(ReadIniValue(CVAPath,"General","Pass",""),chr(9),"")
LOG "Deliverable Name: " & DelivName & " [" & VersionStr & "," & RevisionStr & "," & PassStr & "]"
DelivPN = Replace(ReadIniValue(CVAPath,"General","PN",""),chr(9),"")
LOG "Deliverable PN: " & DelivPN
DelivSize = GetFolderSize(CVAFolder)
LOG "Deliverable Size: " & FormatNumber(DelivSize/1024,0) & " Kbytes"
DriveFreeSize = GetDriveFreeSize("C:\")
LOG "Free Size: " & FormatNumber(DriveFreeSize/1024,0) & " Kbytes"
InstallSize = GetDriveUsedSize("C:\")
LOG "System Used Size: " & FormatNumber(InstallSize/1024,0) & " Kbytes"

if fso.FileExists(CVAFolder & "\src\sysid.txt") AND fso.FileExists("C:\HP\BIN\RStone.INI") then
	Dim SysIdSupp : SysIdSupp = Trimend(ReadAllTextFile(CVAFolder & "\src\sysid.txt"),vbCrLf)
	Dim SysId : SysId = ReadIniValue("C:\HP\BIN\RStone.INI","BIOS Strings","SystemID","")
	if InStr(1,SysIdSupp,SysId,vbTextCompare) = 0 then
		fso.DeleteFile(CVAPath)
		fso.DeleteFile(CVAFolder & "\FAILURE.FLG")
		LOG "System Id Support: No (" & SysIdSupp & ")"
		if fso.FileExists(CVAPath) <> True AND fso.FileExists(CVAFolder & "\FAILURE.FLG") <> True then
			LOG "RESULT=PASSED"
		end if
		SetLocale(sLocale)
		WScript.Quit(0)
	end if
	LOG "System Id Support: Yes (" & SysIdSupp & ")"
end if

FccResult = "FAILED"
Set objShell = WScript.CreateObject("WScript.shell")
LOG "Set process environment variable %FCC_LOG_FOLDER% to """ & LogFolder & """"
set objShellEnv = objShell.Environment("Process")
objShellEnv("FCC_LOG_FOLDER")=LogFolder
objShell.CurrentDirectory = CMDFolder
LOG "Run and no wait install command: " & InstallFile

if fso.FileExists(InstallFile) <> True then LOG "ERROR: Not found [" & InstallFile & "]"


'[2017/10/02] Tony: No wait for process end, log and exit process before FBI timeout
RtnCode = objShell.Run(InstallFile,CONST_HIDE_WINDOW,FALSE)
LOG "Return code: " & RtnCode
LOG "Waiting command finish... " & InstallFile
WaitAndCheck InstallFile, GetTimeoutFromCIA()

objShell.CurrentDirectory = "C:\SWSetup" '[2012/2/17] Watson: need to set folder back else folder delete will fail
Set objShell = Nothing
InstallSize = GetDriveUsedSize("C:\") - InstallSize
LOG "Size Extend: " & FormatNumber(InstallSize/1024,0) & " Kbytes"
if fso.FileExists(CVAFolder & "\FAILURE.FLG") <> True then FccResult = "PASSED"

'[2012/7/9] Watson: Implement metro.xml move to c:\swsetup\metro\deliverable folder
if fso.FileExists(CVAFolder & "\src\Metro.xml") AND FccResult = "PASSED" then
	LOG "Metro Deliverable: Yes"
	Dim MetroFolder : MetroFolder = "C:\SWSetup\Metro\" & DelivPN
	if CreateFolder(MetroFolder) <> True then
		LOG "Metro folder create failed: " & MetroFolder
	else
		fso.CopyFile CVAFolder & "\src\Metro.xml", MetroFolder & "\Metro.xml", True
	end if
else
	LOG "Metro Deliverable: No"
end if

if fso.FileExists("C:\System.sav\Util\Slim.LST") AND FccResult = "PASSED" then
	Dim bSlim : bSlim = false
	for each item in Split(ReadIniSection("C:\System.sav\Util\Slim.LST", "Slim for Installed", ""), vbCrLf)
		if StrComp(DelivName,Trim(item),vbTextCompare) = 0 OR StrComp("*** All Deliverables ***",Trim(item),vbTextCompare) = 0 then
			bSlim = true
			fso.DeleteFolder CVAFolder, True
			if fso.FolderExists(CVAFolder) then 
				LOG "Slim Deliverable: Failed"
				WriteFile CVAFolder, "C:\System.sav\Util\TDC\MCPP\DELETE.LST" '[2012/2/17] Watson: If folder delete failed, add folder path to DELETE.LST
			else
				LOG "Slim Deliverable: for Installed"
			end if
			exit for
		end if
	next
	for each item in Split(ReadIniSection("C:\System.sav\Util\Slim.LST", "Slim for PostLast", ""), vbCrLf)
		if StrComp(DelivName,Trim(item),vbTextCompare) = 0 then
			bSlim = true
			WriteFile CVAFolder, "C:\System.sav\Util\TDC\MCPP\DELETE.LST" '[2012/2/29] Watson: implement to delete folder at PostLast stage to avoid deliverable folder had be removed at 1st install.
			LOG "Slim Deliverable: for PostLast"
			exit for
		end if
	next
	if bSlim = false then LOG "Slim Deliverable: No"
end if
LOG "RESULT=" & FccResult
if fso.FileExists(CVAFolder & "\FAILURE.FLG") then
	WriteFile VbCrLf & "[" & CVAFolder & "\FAILURE.FLG]", LogFile
	WriteFile ReadAllTextFile(CVAFolder & "\FAILURE.FLG"), LogFile
end if

if fso.FileExists("C:\HP\BIN\RStoneFupdate.INI") then 
	InstStage="PASS2"
elseif fso.FileExists("C:\HP\BIN\RStonePre.INI") then
	InstStage="PASS1"
else
	InstStage="UnKnow"
end if

bWrite = WriteIniSection("C:\System.sav\Util\install.bom", "Deliverable List", DelivName & " [" & VersionStr & "," & RevisionStr & "," & PassStr & "]" & ", " & DelivSize & ", " & InstallSize & ", " & DriveFreeSize & ", " & InstStage, "true")

rem IN: C:\System.sav\Logs\[FCC]_%COMPONENT_NAME%\cmdline.txt
rem OUTPUT CONTENT: <UIA code>_<Deliverable Name>: <command line>
Dim CMDLINETXT : CMDLINETXT = LogFolder & "\cmdline.txt"
If fso.FileExists(CMDLINETXT) then
	LOG "Found " & CMDLINETXT & ", write to " & CmdLineTxtSummaryLog
	WriteFile vbCrLf & vbCrLf & "------- " & UIA & "_" & DelivName & " --------------------- " & vbCrLf & ReadAllTextFile(CMDLINETXT) & vbCrLf & "-------------------------------------------------" & vbCrLf & vbCrLf, CmdLineTxtSummaryLog
	LOG "Done, rename to cmdline.summized"
	fso.MoveFile CMDLINETXT, LogFolder & "\cmdline.summized"
Else
	LOG "Not found " & CMDLINETXT
End If


Dim UWPCMD_NAME : UWPCMD_NAME = "appxinst.cmd"
Dim UWPCMD_FLDR : UWPCMD_FLDR = CVAFolder & "\src\uwp"
Dim UWPexCMD : UWPexCMD = LogFolder & "\ex_appxinst.cmd"
Dim UWPexBTO : UWPexBTO = "c:\system.sav\P2PP\BBV2\APPXINST_" & DelivPN & ".BTO"
If fso.FileExists(UWPCMD_FLDR & "\" & UWPCMD_NAME) then
	LOG "Found " & UWPCMD_FLDR & "\" & UWPCMD_NAME & ", create BTO:""" & UWPexBTO & """ and """ & UWPexCMD & """"
	
	'Create exCMD
	' Execution
	WriteFile "echo. >>" & LogFile, UWPexCMD
	WriteFile "echo. >>" & LogFile, UWPexCMD
	WriteFile "pushd """ & UWPCMD_FLDR & """", UWPexCMD
	WriteFile "Echo [%Date% %Time%]" & DelivName & "\appxinst.cmd starts >>" & LogFile, UWPexCMD
	WriteFile "call """ & UWPCMD_NAME & """ >> " & LogFile & " 2>&1" & LogFolder, UWPexCMD
	WriteFile "if errorlevel 1 ( echo RESULT=FAILED>>" & LogFile & " ) else ( echo RESULT=PASSED>>" & LogFile & ")", UWPexCMD
	WriteFile "Echo [%Date% %Time%]" & DelivName & "\appxinst.cmd end >>" & LogFile, UWPexCMD
	WriteFile "popd", UWPexCMD
	WriteFile "echo. >>" & LogFile, UWPexCMD
	' Print cmdline.txt
	WriteFile "echo. >>" & CmdLineTxtSummaryLog, UWPexCMD
	WriteFile "echo. >>" & CmdLineTxtSummaryLog, UWPexCMD
	WriteFile "echo ------- UWP of " & DelivName & " --------------------- >> """ & CmdLineTxtSummaryLog & """", UWPexCMD
	WriteFile "if exist """ & LogFolder & "\cmdline.txt"" type """ & LogFolder & "\cmdline.txt"" >> """ & CmdLineTxtSummaryLog & """", UWPexCMD
	WriteFile "echo. >>" & CmdLineTxtSummaryLog, UWPexCMD
	WriteFile "echo. >>" & CmdLineTxtSummaryLog, UWPexCMD
	
	'Create exBTO
	WriteFile "@SingleObj cmd.exe /c """ & UWPexCMD & """ " & LogFolder, UWPexBTO
	LOG "Done"
Else
	LOG "Not found " & UWPCMD_FLDR & "\" & UWPCMD_NAME
End If




SetLocale(sLocale)
WScript.Quit(0)




















Function ReadAllTextFile(filename)
	Const ForReading = 1, ForWriting = 2
	Dim fso, objfile
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objfile = fso.OpenTextFile(filename, ForReading)

	If objfile.AtEndOfStream Then
		ReadAllTextFile = ""
	Else
		ReadAllTextFile = objfile.ReadAll
	End If
	objfile.Close
	set objfile = Nothing
End Function

Function FindString(findstr, filename)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim bFound : bFound = False
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim objfile : Set objfile = fso.OpenTextFile(filename, ForReading)
	Do Until objFile.AtEndOfStream
		strSearchString = objFile.ReadLine
		if StrComp(strSearchString, findstr, vbTextCompare) = 0 then
			bFound = True
			Exit Do
		end if
	Loop
	objfile.Close
	set objfile = Nothing
	FindString = bFound
End Function

Function GetFileExtend(InstallFolder, extendstr)
   Dim fso, folder, file, fc
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set folder = fso.GetFolder(InstallFolder)
   Set fc = folder.Files
   For Each file in fc
	  'if InStr(Len(file.name)-3,file.name, extendstr, vbTextCompare) <> 0 then
	  if StrComp(fso.GetExtensionName(file.name), extendstr, vbTextCompare) = 0 then
		 GetFileExtend = file.name
	  end if
   Next
   Set folder = Nothing
End Function

Function CreateFolder(path)
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim ParentFolder : ParentFolder = fso.GetParentFolderName(path)
	if fso.FolderExists(ParentFolder) <> True AND Len(ParentFolder) <> 0 then CreateFolder(ParentFolder)
	if fso.FolderExists(path) <> True then fso.CreateFolder(path)
	CreateFolder = fso.FolderExists(path)
End Function

Function ReadIniValue(inifile,section,key,default)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
   Dim fso, objFile, strText, strSection, strValue, PosSection, PosEndSection, PosValue, PosEndValue
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set objFile = fso.OpenTextFile(inifile, ForReading, False, TristateUseDefault)
   strText = objFile.ReadAll
   objFile.Close
   set objFile = Nothing
   
   strValue = default
   'Find section
   PosSection = InStr(1, strText, "[" & section & "]", vbBinaryCompare)
   If PosSection>0 Then
	  'Section exists. Find end of section
	  PosEndSection = InStr(PosSection, strText, vbCrLf & "[")
	  '?Is this last section?
	  If PosEndSection = 0 Then PosEndSection = Len(strText)+1
	  'Separate section contents
	  strSection = Mid(strText, PosSection, PosEndSection - PosSection)
	  strSection = split(strSection, vbCrLf)
	  key = key & "="
	  For Each Line In strSection
		 If StrComp(Left(Line, Len(key)), key, vbTextCompare) = 0 Then
			strValue = Mid(Line, Len(key)+1)
		 End If
	  Next
   End If
   ReadIniValue = strValue
End Function


Function ReadIniSection(inifile,section,default)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
   Dim fso, objFile, strText, strSection, PosSection, PosEndSection
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set objFile = fso.OpenTextFile(inifile, ForReading, False, TristateUseDefault)
   strText = objFile.ReadAll
   objFile.Close
   set objFile = Nothing
   
   strSection = default
   'Find section
   PosSection = InStr(1, strText, "[" & section & "]", vbBinaryCompare)
   If PosSection>0 Then
	  'Section exists. Find end of section
	  PosEndSection = InStr(PosSection, strText, vbCrLf & "[")
	  '?Is this last section?
	  If PosEndSection = 0 Then PosEndSection = Len(strText)+1
	  'Separate section contents
	  strSection = Mid(strText, PosSection, PosEndSection - PosSection)
   End If
   ReadIniSection = strSection
End Function


Function WriteIniSection(inifile,section,content,append)
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Dim fso, objFile, strText, strSection, strAfter, PosSection, PosEndSection
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set objFile = fso.OpenTextFile(inifile, ForReading, True)
   If objFile.AtEndOfStream Then
		strText = ""
   Else
		strText = objFile.ReadAll
   End If
   objFile.Close
   set objFile = Nothing
   
   PosSection = InStr(1, strText, "[" & section & "]", vbTextCompare)
   If PosSection>0 Then
	  'Section exists. Find end of section
	  PosEndSection = InStr(PosSection, strText, vbCrLf & "[")
	  '?Is this last section?
	  If PosEndSection = 0 Then PosEndSection = Len(strText)+1
	  do while Mid(strText,PosEndSection-2,2) = vbCrLf
		 PosEndSection=PosEndSection-2
	  Loop
	  strSection = Mid(strText, PosSection, PosEndSection - PosSection)
	  If StrComp(append, "true", vbTextCompare) = 0 Then
		 If Right(strSection,2) <> vbCrLf Then strSection = strSection & vbCrLf
		 strSection = strSection & content
	  Else
		 strSection = Left(strSection, Len(section)+4) & content & Right(strSection,Len(strSection)-Len(section)-2)
	  End If
	  strAfter = Left(strText, PosSection-1) & strSection & Right(strText,Len(strText)-(PosEndSection-1))
	  Set objFile = fso.OpenTextFile(inifile, ForWriting)
	  objFile.Write strAfter
	  objFile.Close
	  set objFile = Nothing
   Else
	  strSection = "[" & section & "]" & vbCrLf & content
	  If Right(strText,2)=vbCrLf or strText="" Then 
		 strAfter = strText & strSection
	  Else
		 strAfter = strText & vbCrLf & strSection
	  End If
	  Set objFile = fso.OpenTextFile(inifile, ForWriting, True)
	  objFile.Write strAfter
	  objFile.Close
	  set objFile = Nothing
   End If

   WriteIniSection = 0
End Function

Sub LOG(msg)
	WriteFile "[" & FormatDateTime(Time, 3) & "] " & msg, LogFile
End Sub




Function WriteFile(strOut,file)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim objFs : Set objFs = CreateObject("Scripting.FileSystemObject")
	Dim objFile : Set objFile = objFs.OpenTextFile(file, ForAppending, True)
	If IsObject(objFile) <> True Then Exit Function
	objFile.Write strOut & VbCrLf
	objFile.Close
	Set objFile = Nothing
End Function


Function Trimend(word, trimchar)
	Dim newword : newword = word
	if Right(newword,1) = trimchar then
		newword = Left(newword, Len(word) - Len(trimchar))
		newword = Trimend(newword, trimchar)
	end if 
	Trimend = newword
End Function


Function GetFolderSize(path)
   Dim fso, folder
   On Error Resume Next
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set folder = fso.GetFolder(path)
   GetFolderSize = folder.size
   Set folder = Nothing
   On Error goto 0
End Function


Function GetDriveUsedSize(drvpath)
   Dim fso, d
   On Error Resume Next
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(drvpath)))
   GetDriveUsedSize = d.TotalSize - d.AvailableSpace
   On Error goto 0
End Function


Function GetDriveFreeSize(drvpath)
   Dim fso, d
   On Error Resume Next
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set d = fso.GetDrive(fso.GetDriveName(fso.GetAbsolutePathName(drvpath)))
   GetDriveFreeSize = d.FreeSpace
   On Error goto 0
End Function


'[Configuration]
'Timeout=1800;C:\SYSTEM.SAV\FBI\STATE.INI
'[FBITB.SpawnProgram]
'Timeout=7200
Function GetTimeoutFromCIA()
	GetTimeoutFromCIA = defaultTimeoutMaxSeconds
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	LOG "Getting timeout value from CIA"
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim objCia : Set objCia = fso.OpenTextFile("c:\SYSTEM.SAV\Util\CIA.INI", ForReading, False, TristateUseDefault)
	Dim iniCia : iniCia = objCia.ReadAll
	objCia.Close
	set objCia = Nothing
	Dim strTmp : strTmp = ReadIniValue_inibuffer(iniCia,"Configuration","Timeout","900")
	'LOG "strTmp=" & strTmp
	If InStr(1, strTmp, ";") > 1 then strTmp = Left(strTmp, InStr(1, strTmp, ";")-1)
	GetTimeoutFromCIA = CInt(strTmp)
	LOG "GetTimeoutFromCIA=" & GetTimeoutFromCIA
End Function


Function ReadIniValue_inibuffer(strText,section,key,default)
	Dim strSection, strValue, Line, PosSection, PosEndSection
	strValue = default
	'Find section
	PosSection = InStr(1, strText, "[" & section & "]", vbBinaryCompare)
	If PosSection>0 Then
		'Section exists. Find end of section
		PosEndSection = InStr(PosSection, strText, vbCrLf & "[")
		'?Is this last section?
		If PosEndSection = 0 Then PosEndSection = Len(strText)+1
		'Separate section contents
		strSection = Mid(strText, PosSection, PosEndSection - PosSection)
		strSection = split(strSection, vbCrLf)
		key = key & "="
		For Each Line In strSection
			If StrComp(Left(Line, Len(key)), key, vbTextCompare) = 0 Then
				strValue = Mid(Line, Len(key)+1)
			End If
		Next
	End If
	ReadIniValue_inibuffer = Trim(strValue)
End Function





Sub WaitAndCheck(installCmd, timeoutMaxSec)
	LOG "Max wait timeout: " & timeoutMaxSec & " seconds"
	Dim secMax : secMax = timeoutMaxSec - (secStep * 2)
	Dim secWait, procStatus
	For secWait = 0 To secMax Step secStep
		procStatus = ChkProcess(installCmd)
		If procStatus = 0 then
			LOG "[" & secWait & "] Process is still running, wait another " & secStep & " seconds"
			Wscript.Sleep (secStep * 1000)
		ElseIf procStatus < 0 then
			LOG "Process is finished"
			Exit Sub
		End If
	Next
	LOG "Time's up, component timeout, process will terminated by FBI soon"
End Sub

'Return -1: process is not found
'Return  0: process is     found
Function ChkProcess(installCmd)
	ChkProcess = -1
	Dim colProcs, objProc
	Set colProcs = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery("Select * From Win32_Process Where Caption = ""cmd.exe""")
	'Set colProcs = CreateObject("WbemScripting.SWbemLocator").ConnectServer.ExecQuery("Select * From Win32_Process")
	Dim posToken1, posToken2
	For Each objProc in colProcs
		'LOG "objProc information: {Caption=" & objProc.Caption & " CommandLine=" & objProc.CommandLine & "}"
		If InStr(1, objProc.CommandLine, installCmd, vbTextCompare) > 1 OR InStr(1, objProc.CommandLine, Replace(installCmd, ".\", "\"), vbTextCompare) > 1 Then 
			If InStr(1, objProc.CommandLine, "cscript.exe", vbTextCompare) <= 0 Then
				ChkProcess = 0
				'LOG "Process is running"
				Exit Function
			End If
		End If
	Next
End Function

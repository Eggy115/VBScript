Option Explicit

If Wscript.Arguments.Count <> 2 Then
   Help()
   WSCript.Quit 1
End If

Dim cva_path : cva_path = Wscript.Arguments(0)
Dim cmd_path : cmd_path = Wscript.Arguments(1)

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
if not fso.FileExists(cva_path) then
	WScript.Echo "ERROR: cva not found."
	Help()
	WScript.Quit(1)
end if


if fso.FileExists(cmd_path) then fso.DeleteFile cmd_path

LOG "Software Title=" & ReadIniValue(cva_path, "Software Title", "US", "")
LOG "Vender Version=" & ReadIniValue(cva_path, "General", "VendorVersion", "")
LOG "PN=" & ReadIniValue(cva_path, "General", "PN", "")

'[Install Execution]
'Install=HPRBWINx64.EXE /CPO:$PATH$
'SilentInstall="install.cmd"
Dim sUtilCmd : sUtilCmd = ReadIniValue(cva_path, "Install Execution", "Install", "")
LOG "sUtilCmd=[" & sUtilCmd & "]"
Dim sUtilCmd_ex : sUtilCmd_ex = Replace(sUtilCmd, "$PATH$" , "%~1")
LOG "sUtilCmd_ex=[" & sUtilCmd_ex & "]"
WriteFile "@echo off",cmd_path
WriteFile "set errorlevel=",cmd_path
WriteFile sUtilCmd_ex,cmd_path
WriteFile "echo errorlevel=%errorlevel%", cmd_path

'[ReturnCode]
'0:SUCCESS:NOREBOOT=No error
'3010:SUCCESS:NOREBOOT=No error
'224:FAILURE:NOREBOOT=Can not find BID
'1:FAILURE:NOREBOOT=Command Typo
'16:FAILURE:NOREBOOT=Driver issue, Include unable to load driver and driver version is too old
'50:FAILURE:NOREBOOT=Unable to open file
'34:FAILURE:NOREBOOT=Fail to allocate memory
'70:FAILURE:NOREBOOT=Fail to get BIOS info
'77:FAILURE:NOREBOOT=Get block size error
Dim lstReturnCode : lstReturnCode = ReadIniSection_woTitle(cva_path,"ReturnCode","")
LOG "lstReturnCode=[" & lstReturnCode & "]"
Dim line
For Each line In split(lstReturnCode, vbCrLf)
	line=Trim(line)
	if Len(line)>0 then
		if isEcPass(line) then 
			WriteFile "if %errorlevel% EQU " & getEcCode(line) & " exit /b 0", cmd_path
		else
			WriteFile "if %errorlevel% EQU " & getEcCode(line) & " echo " & getEcMsg(line) & " & exit /b %errorlevel%", cmd_path
		end if
	end if
Next
WriteFile "echo CVA not define return code %errorlevel% & exit /b 1", cmd_path
WSCript.Quit 0


Sub LOG(msg)
	WScript.echo msg
End Sub

Function isEcPass(sLine)
	isEcPass = false
	if InStr(1, sLine, ":SUCCESS:", vbTextCompare)>0 then isEcPass = true
end function

Function getEcCode(sLine)
	getEcCode = 0
	Dim posE : posE = InStr(1, sLine, ":", vbTextCompare) - 1
	if posE > 0 then getEcCode = CInt(Mid(sLine, 1, posE))
End Function

Function getEcMsg(sLine)
	getEcMsg = ""
	Dim posB : posB = InStrRev(sLine, "=", -1, vbTextCompare)
	if posB > 0 then getEcMsg = Mid(sLine, posB+1)
End Function


Function WriteFile(strOut,file)
	'WScript.Echo "WriteFile(" & strOut & "," & file & ")"
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim objFs : Set objFs = CreateObject("Scripting.FileSystemObject")
	Dim objFile : Set objFile = objFs.OpenTextFile(file, ForAppending, True)
	If IsObject(objFile) <> True Then Exit Function
	objFile.Write strOut & VbCrLf
	objFile.Close
	Set objFile = Nothing
End Function

Function ReadIniSection_woTitle(inifile,section,default)
	ReadIniSection_woTitle = ReadIniSection(inifile,section,default)
	Dim posB : posB = InStr(1, ReadIniSection_woTitle, "[" & section & "]")
	if posB <=0 then Exit Function

	posB = InStr(posB, ReadIniSection_woTitle, vbCrLf)
	if posB <=0 OR posB+Len(vbCrLf) > Len(ReadIniSection_woTitle) then
		ReadIniSection_woTitle = ""
		Exit Function
	end if
	ReadIniSection_woTitle = Mid(ReadIniSection_woTitle, posB+Len(vbCrLf))
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

Function ReadIniValue(inifile,section,key,default)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Dim fso, objFile, strText, strSection, strValue, PosSection, PosEndSection, PosValue, PosEndValue, Line
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


Sub Help()
	WScript.Echo "Syntax:" & chr(13) & chr(13) & "GenBiosReadBackScript.VBS <IN:CVA_FULL_PATH> <OUT:INSTALLCMD_PATH>"
	WSCript.Quit(1)
End Sub

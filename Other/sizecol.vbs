Option Explicit
On Error Resume Next

const scriptVer="1.00,A,1 rev 100"
'   1.00,A,1 rev 100
'   tony.wu@hp.com
'   Initial release

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim LogFile : LogFile = "c:\system.sav\logs\ImgEnh\sizecol.log"
Dim strCreateFolder
strCreateFolder = CreateFolder("c:\system.sav\logs\ImgEnh")

Dim sLocale : sLocale = "en-us"
sLocale = getLocale()
SetLocale("en-us")
LOG vbCrLf & vbCrLf & "===== SizeCol.vbs [" & scriptVer & "] ====================================================================="
LOG "Current Locale: " & sLocale

Dim InstallBom : InstallBom = "C:\System.sav\Util\Install.bom"
strCreateFolder = CreateFolder("c:\system.sav\Util")

' Columns to be in the report
Dim DelivName : DelivName = "Undefined"
Dim VersionStr : VersionStr = "-.--"
Dim RevisionStr : RevisionStr = "-"
Dim PassStr : PassStr = "-"
Dim DelivSize : DelivSize = "NA"
Dim DriveFreeSize
Dim InstallSize
Dim InstStage : InstStage = "UnKnow"
LOG "Version and DelivSize are fixed string"


' --- DelivName ----------------------------------------------
DelivName = Trim(Wscript.Arguments(0))
LOG "DelivName: """ & DelivName & """"

' --- DriveFreeSize ----------------------------------------------
DriveFreeSize = GetDriveFreeSize("C:\")
LOG "Free Size: " & FormatNumber(DriveFreeSize/1024,0) & " Kbytes"

' --- InstallSize ----------------------------------------------
InstallSize = GetDriveUsedSize("C:\")
LOG "System Used Size: " & FormatNumber(InstallSize/1024,0) & " Kbytes"

' --- InstStage ----------------------------------------------
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
if fso.FileExists("C:\HP\BIN\RStoneFupdate.INI") then 
	InstStage="PASS2"
elseif fso.FileExists("C:\HP\BIN\RStonePre.INI") then
	InstStage="PASS1"
end if
LOG "InstStage: """ & InstStage & """"

' --- Write to install.bom ----------------------------------------------
LOG "Write to install.bom"
Dim bWrite : bWrite = WriteIniSection(InstallBom, "Deliverable List", DelivName & " [" & VersionStr & "," & RevisionStr & "," & PassStr & "]" & ", " & DelivSize & ", " & InstallSize & ", " & DriveFreeSize & ", " & InstStage, "true")
LOG "Result: " & bWrite

set fso = nothing
SetLocale(sLocale)
WScript.Quit(0)






' === Sub functions ============================================================================================================================================


Function CreateFolder(path)
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim ParentFolder : ParentFolder = fso.GetParentFolderName(path)
	if fso.FolderExists(ParentFolder) <> True AND Len(ParentFolder) <> 0 then CreateFolder(ParentFolder)
	if fso.FolderExists(path) <> True then fso.CreateFolder(path)
	CreateFolder = fso.FolderExists(path)
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









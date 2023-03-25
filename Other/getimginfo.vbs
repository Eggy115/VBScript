'On Error Resume Next
DIM DEBUG : DEBUG = TRUE

'v3.00,A,8   2021/06/07
'	Tony Wu
'	Update item #1: Support Pro
'
'v3.00,A,7   2021/06/04
'	Tony Wu
'	Update item #1: Support W11 upgrade / downloadgrade
'
'v3.00,A,6   2020/09/03
'	Tony Wu
'	Update item #1: AK4=AC4
'
'v3.00,A,5   2018/06/29
'	Tony Wu
'	Update item #1: Supports AddOptMap
'
'v3.00,A,4   2018/06/12
'	Tony Wu
'	Update item #1: Fix WKS+DG OS flag checking

'v3.00,A,3   2018/05/29
'	Tony Wu
'	Update item #1: Never block between EDU and non-EDU, WKS and non-WKS
'	Update item #2: Log path change from c:\system.sav\logs\BB\ImgEnh\getimginfo.log to c:\system.sav\logs\ImgEnh\getimginfo.log


CONST CONST_WIN7SP1_VERSION = "6.1.7601"
CONST CONST_WIN8_VERSION    = "6.2.9200"
CONST CONST_WIN81_VERSION   = "6.3.9200"
CONST CONST_WIN81U_VERSION  = "6.3.9600"
CONST CONST_WIN10_VERSION   = "6.4"


Dim CurrentFolder : CurrentFolder = "NA"
Dim IdPath        : IdPath        = "NA"
Dim FlagPath      : FlagPath      = "NA"
Dim RStoneUUTPath : RStoneUUTPath = "NA"
Dim RStone1GSPath : RStone1GSPath = "NA"
Dim DmiPath       : DmiPath       = "NA"
Dim AomPath       : AomPath       = "NA"



Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
CurrentFolder = fso.GetParentFolderName(WScript.ScriptFullName)

Dim LogFile : LogFile = "x:\getimginfo.log"
if fso.FolderExists("c:\system.sav\logs") then 
	LogFile = "c:\system.sav\logs\ImgEnh\getimginfo.log"
	WScript.Echo "LogFile=" & LogFile
	MakeDirByFilePath LogFile
end if
LOG vbCrLf & vbCrLf & vbCrLf & "=== A new Log of GetImgInfo.vbs start ============================================================"

if fso.FileExists(  CurrentFolder & "\RStone.ini")    then RStoneUUTPath = CurrentFolder & "\RStone.ini"    else RStoneUUTPath = "C:\HP\BIN\RStone.ini"
if fso.FileExists(  CurrentFolder & "\RStone1GS.ini") then RStone1GSPath = CurrentFolder & "\RStone1GS.ini" else RStone1GSPath = "C:\HP\BIN\RStone1GS.ini"
if fso.FileExists(  CurrentFolder & "\sysid.txt")     then IdPath        = CurrentFolder & "\sysid.txt"     else IdPath        = "C:\System.sav\util\sysid.txt"
if fso.FolderExists(CurrentFolder & "\Flags")         then FlagPath      = CurrentFolder & "\Flags"         else FlagPath      = "C:\System.sav\Flags"

if fso.FileExists(CurrentFolder & "\DMI.ini") then 
	DmiPath = CurrentFolder & "\DMI.ini"
elseif fso.FileExists("C:\SYSTEM.SAV\TWEAKS\DMI.ini") then 
	DmiPath = "C:\SYSTEM.SAV\TWEAKS\DMI.ini"
elseif fso.FolderExists("C:\SYSTEM.SAV\TWEAKS\DMI") then 
	DmiPath = FindDmiIniPath("C:\SYSTEM.SAV\TWEAKS\DMI")
end if
LOG "DmiPath=""" & DmiPath & """"

if fso.FileExists(CurrentFolder & "\AddOptMap.ini") then 
	AomPath = CurrentFolder & "\AddOptMap.ini"
elseif fso.FileExists("C:\HP\Support\AddOptMap.ini") then 
	AomPath =   "C:\HP\Support\AddOptMap.ini"
end if
LOG "AomPath=""" & AomPath & """"

if StrComp(DmiPath, "NA", vbTextCompare) = 0 AND StrComp(AomPath, "NA", vbTextCompare) = 0 then
	ERR "Found neither DMI.ini nor AddOptMap.ini file."
	ERR "Fail=1"
	WScript.Quit (1)
end if
if fso.FileExists(RStoneUUTPath) <> True then
	ERR "Could not found RStone.ini file."
	ERR "Fail=1"
	WScript.Quit (1)
end if
if fso.FileExists(RStone1GSPath) <> True then
	ERR "Could not found RStone1GS.ini file."
	ERR "Fail=1"
	WScript.Quit (1)
end if
if fso.FileExists(IdPath) <> True then
	ERR "Could not found sysid.txt file."
	ERR "Fail=1"
	WScript.Quit (1)
end if

LOG "RStoneUUTPath=" & RStoneUUTPath
LOG "RStone1GSPath=" & RStone1GSPath
LOG "IdPath=" & IdPath
LOG "FlagPath=" & FlagPath

'About FB
Dim FB_IMG    : FB_IMG    = ReadIniValue(RStone1GSPath,"BIOS Strings","FeatureByte",     "0")
Dim FB_UUT    : FB_UUT    = ReadIniValue(RStoneUUTPath,"BIOS Strings","FeatureByte",     "0")
LOG "FB_UUT=""" & FB_UUT & """"
LOG "FB_IMG=""" & FB_IMG & """"

'About OEM Activation and Downgrade Program
Dim ImgIDFlag : ImgIDFlag = ReadIniValue(RStoneUUTPath,"BIOS Strings","ImageIDFlag",     "0") 'ImageIDFlag check W7Downgrade(Windows 7 Downgrade (ATF_OS.7,Vos.B,OA3) [6747aq=4]
Dim OA21Flag  : OA21Flag  = ReadIniValue(RStoneUUTPath,"BIOS Strings","OEMActivation21", "0")
Dim OA30Flag  : OA30Flag  = ReadIniValue(RStoneUUTPath,"BIOS Strings","OEMActivation30", "0")
LOG "ImgIDFlag=""" & ImgIDFlag & """"
LOG "OA21Flag=""" & OA21Flag & """"
LOG "OA30Flag=""" & OA30Flag & """"

'About SSID
Dim SysId     : SysId     = ReadIniValue(RStoneUUTPath,"BIOS Strings","SystemID", "0")
Dim SysIdList : SysIdList = Trimend(ReadAllTextFile(IdPath),vbCrLf)
LOG "SystemId=" & SysId
LOG "SystemIdList=" & SysIdList

'About edition servicing
Dim OSSku_UUT     : OSSku_UUT     = ReadIniValue(RStoneUUTPath,"BIOS Strings","OSSkuFlag",           "0")
Dim OSSku_IMG     : OSSku_IMG     = ReadIniValue(RStone1GSPath,"BIOS Strings","OSSkuFlag",           "0")
Dim OSEdition_UUT : OSEdition_UUT = ReadIniValue(RStoneUUTPath,"BIOS Strings","WinConfigurationFlag","0")
Dim OSEdition_IMG : OSEdition_IMG = ReadIniValue(RStone1GSPath,"BIOS Strings","WinConfigurationFlag","0")
LOG "OSSku_UUT=""" & OSSku_UUT & """"
LOG "OSSku_IMG=""" & OSSku_IMG & """"
LOG "OSEdition_UUT=""" & OSEdition_UUT & """"
LOG "OSEdition_IMG=""" & OSEdition_IMG & """"

'About location
Dim OptionCodeS_IMG : OptionCodeS_IMG = ReadIniValue(RStone1GSPath,"BIOS Strings","SKUOptionConfig", "0")
Dim OptionCodeS_UUT : OptionCodeS_UUT = ReadIniValue(RStoneUUTPath,"BIOS Strings","SKUOptionConfig", "0")
LOG "OptionCodeS_UUT=""" & OptionCodeS_UUT & """"
LOG "OptionCodeS_IMG=""" & OptionCodeS_IMG & """"

'About Compact/nonCompact
Dim isCompact_IMG : isCompact_IMG = CInt(ReadIniValue(RStone1GSPath,"BIOS Strings","Compact", "0"))
Dim isCompact_UUT : isCompact_UUT = CInt(ReadIniValue(RStoneUUTPath,"BIOS Strings","Compact", "0"))
LOG "isCompact_UUT=""" & isCompact_UUT & """"
LOG "isCompact_IMG=""" & isCompact_IMG & """"

'About S_MODE/CloudOS
'RS3:Vos.CL + ATF_OS.C(gJ)
'RS4:S_MODE(hY)
Dim isCloudOS_IMG : isCloudOS_IMG = 0
if isCloudOS_IMG<>1 then isCloudOS_IMG = CInt(ReadIniValue(RStone1GSPath,"BIOS Strings","CloudOS", "0"))
if isCloudOS_IMG<>1 then isCloudOS_IMG = CInt(ReadIniValue(RStone1GSPath,"BIOS Strings","S_MODE",  "0"))
Dim isCloudOS_UUT : isCloudOS_UUT = 0
if isCloudOS_UUT<>1 then isCloudOS_UUT = CInt(ReadIniValue(RStoneUUTPath,"BIOS Strings","CloudOS", "0"))
if isCloudOS_UUT<>1 then isCloudOS_UUT = CInt(ReadIniValue(RStoneUUTPath,"BIOS Strings","S_MODE",  "0"))
LOG "isCloudOS_UUT=""" & isCloudOS_UUT & """"
LOG "isCloudOS_IMG=""" & isCloudOS_IMG & """"

'About Chassis Feature Byte
Dim ChassisFB : ChassisFB = "NA"
If InStr(1,SysIdList,"#",vbTextCompare) <> 0 Then 'CDT
	If StrComp(ChassisFB, "NA", vbTextCompare) =  0 AND StrComp(AomPath,   "NA", vbTextCompare) <> 0 Then 
		LOG "Get chassis FB from AddOptMap.ini"
		ChassisFB = ParseFeatureByte(FB_UUT, AomPath, "Option_Map", "C_")
	End If
	If StrComp(ChassisFB, "NA", vbTextCompare) =  0 AND StrComp(DmiPath,   "NA", vbTextCompare) <> 0 Then  
		LOG "Get chassis FB from DMI.ini"
		ChassisFB = ParseFeatureByte(FB_UUT, DmiPath, "Options",    "C_")
	End If
end if
LOG "ChassisFB=""" & ChassisFB & """"




rem ----- Check Downgrade Program ----------------------------------------------------------------------


Dim ImgSku : ImgSku = "Unknow"
'Remove FB_UUT = "0" cases since it's Legacy not Fusion
'ImgIDFlag=4   means Win7 support OA3
'isDGOS = TRUE means Workstation(Vos.WKS) or Education(Vos.STU) or Professional(Vos.B)
Dim isDGOS : isDGOS=false
if OSEdition_UUT = 3 OR OSEdition_UUT = 7 OR OSEdition_UUT = 9 then isDGOS = true

rem _DG mean the UUT has ability to downgrade to
rem For example, Win7_DG means the unit has the ability to downgrade to Win7 (shipping OS may be Win10)
if     OSSku_UUT = 2 AND ImgIDFlag = 0 AND OA21Flag = 1 AND OA30Flag = 0 then
	ImgSku = "Win7_STD"
elseif OSSku_UUT = 2 AND ImgIDFlag = 0 AND OA21Flag = 1 AND OA30Flag = 1 AND isDGOS then
	ImgSku = "Win7_DG"
elseif OSSku_UUT = 2 AND ImgIDFlag = 4 AND OA21Flag = 1 AND OA30Flag = 1 AND isDGOS then
	ImgSku = "Win7_DG"
elseif OSSku_UUT = 5 AND ImgIDFlag = 0 AND OA30Flag = 1 AND not isDGOS then
	ImgSku = "Win10"
elseif OSSku_UUT = 5 AND ImgIDFlag = 0 AND OA30Flag = 1 AND isDGOS then
	ImgSku = "Win10_DG"
elseif OSSku_UUT = 6 AND ImgIDFlag = 0 AND OA30Flag = 1 AND not isDGOS then
	ImgSku = "Win11"
elseif OSSku_UUT = 6 AND ImgIDFlag = 0 AND OA30Flag = 1 AND isDGOS then
	ImgSku = "Win11_DG"
end if
LOG "ImgSku=""" & ImgSku & """"


rem ----- Check flags ----------------------------------------------------------------------
Dim OSFlagChk : OSFlagChk = 0
Select Case ImgSku
Case "Win7_STD"
	if fso.FileExists(FlagPath & "\Win7sp1.flg") = True Then OSFlagChk = 1
Case "Win7_DG"
	if fso.FileExists(FlagPath & "\Win7sp1.flg") = True Then OSFlagChk = 1
	if fso.FileExists(FlagPath & "\w10.flg")     = True Then OSFlagChk = 1
Case "Win10"
	rem DASH:    Win10 from ATF_OS.T as non-NPI platforms go Win Oct21
	rem SSRM:    Win10 from ATF_OS.T as non-NPI platforms go Win 10/Oct21, can       upgrade to Win Oct21?
	if fso.FileExists(FlagPath & "\W10.flg") = True Then OSFlagChk = 1
	if fso.FileExists(FlagPath & "\W11.flg") = True Then OSFlagChk = 1
Case "Win10_DG"
	if fso.FileExists(FlagPath & "\W10.flg") = True Then OSFlagChk = 1
	if fso.FileExists(FlagPath & "\W11.flg") = True Then OSFlagChk = 1
Case "Win11"
	rem DASH:    Win11 from ATF_OS.Z as     NPI platforms go Win Oct21
	rem SSRM:    Win11 from ATF_OS.Z as     NPI platforms go Win    Oct21, can downloadgrade to Win 10
	if fso.FileExists(FlagPath & "\W11.flg") = True Then OSFlagChk = 1
	if fso.FileExists(FlagPath & "\W10.flg") = True Then OSFlagChk = 1
Case "Win11_DG"
	if fso.FileExists(FlagPath & "\W11.flg") = True Then OSFlagChk = 1
	if fso.FileExists(FlagPath & "\W10.flg") = True Then OSFlagChk = 1
End Select
LOG ">>>>>>> OSFlagChk=" & OSFlagChk


rem ----- Check EditionServicing ----------------------------------------------------------------------
Dim EditionServicingChk : EditionServicingChk = 1
'--- 3.00,A,3 ---------------------------------------
LOG "Update in 3.00,A,3, never check edition, force set EditionServicingChk as 1"
EditionServicingChk = 1
'--- 3.00,A,1 ---------------------------------------
rem				'Dim OSES_UUT : OSES_UUT = "" & OSEdition_UUT
rem				'Dim OSES_IMG : OSES_IMG = "" & OSEdition_IMG
rem				Dim OSES_UUT : OSES_UUT = "SKIP_CHECKING"
rem				Dim OSES_IMG : OSES_IMG = "SKIP_CHECKING"
rem				LOG "OSES_IMG=" & OSES_IMG
rem				LOG "OSES_UUT=" & OSES_UUT
rem				'RS3 Workstartion: VOS.B   + DPK_WKS(fj) --> WinConfigurationFlag=3
rem				'RS4 Workstartion: VOS.WKS + DPK_WKS(fj) --> WinConfigurationFlag=9
rem				if FB_Search(FB_IMG, "fj") And ( OSEdition_IMG=3 OR OSEdition_IMG=9 ) then OSES_IMG = "Win10-WKS"
rem				if FB_Search(FB_UUT, "fj") And ( OSEdition_UUT=3 OR OSEdition_UUT=9 ) then OSES_UUT = "Win10-WKS"
rem				'RS3 Education: VOS.B   + NaAc(aQ)       --> WinConfigurationFlag=3
rem				'RS4 Education: VOS.STU + NaAc(aQ)       --> WinConfigurationFlag=7
rem				if FB_Search(FB_IMG, "aQ") And ( OSEdition_IMG=3 OR OSEdition_IMG=7 ) then OSES_IMG = "Win10-EDU"
rem				if FB_Search(FB_UUT, "aQ") And ( OSEdition_UUT=3 OR OSEdition_UUT=7 ) then OSES_UUT = "Win10-EDU"
rem				LOG "OSES_IMG=" & OSES_IMG
rem				LOG "OSES_UUT=" & OSES_UUT
rem				If StrComp(OSES_IMG, OSES_UUT, vbTextCompare) <> 0 then EditionServicingChk = 0
LOG ">>>>>>> EditionServicingChk=" & EditionServicingChk


rem ----- Check OptionCode ----------------------------------------------------------------------
Dim OptionCodeChk : OptionCodeChk = 1
rem 2020/09/03 Potential risk on the customer units with #AC4 and later restore #AK4 image/SSRM will cause failure by system lock. [Tony] will update syslock component to allow both #AK4 and #AC4 can successfully restore.
if StrComp(OptionCodeS_IMG, "AK4", vbTextCompare)=0 Then OptionCodeS_IMG = "Brazil(AK4/AC4)"
if StrComp(OptionCodeS_IMG, "AC4", vbTextCompare)=0 Then OptionCodeS_IMG = "Brazil(AK4/AC4)"
if StrComp(OptionCodeS_UUT, "AK4", vbTextCompare)=0 Then OptionCodeS_UUT = "Brazil(AK4/AC4)"
if StrComp(OptionCodeS_UUT, "AC4", vbTextCompare)=0 Then OptionCodeS_UUT = "Brazil(AK4/AC4)"
if StrComp(OptionCodeS_IMG, OptionCodeS_UUT, vbTextCompare)<>0 Then OptionCodeChk = 0
LOG ">>>>>>> OptionCodeCheck=" & OptionCodeChk


rem ----- Check Compact and CloudOS ----------------------------------------------------------------------
Dim ExtraChk : ExtraChk = 1
if isCompact_IMG <> isCompact_UUT Then ExtraChk = 0
if isCloudOS_IMG <> isCloudOS_UUT Then ExtraChk = 0
LOG ">>>>>>> ExtraFeatureCheck=" & ExtraChk


rem ----- Check System Board ID and Chassis Feature Byte ----------------------------------------------------------------------
Dim SysIDChk : SysIDChk = 0
If InStr(1,SysIdList,"#",vbTextCompare) <> 0 then
	REM (CDT) 81B7#C_Fan27,81B9#C_Fan23,81B8#C_Fan23,81B7#C_Fan23,81B7#C_FanX23,81BA#C_PRO195,81BB#C_PRO195,81BC#C_PRO195,81BA#C_PRO215,81BB#C_PRO215,81BC#C_PRO215,81BA#C_PRO238,81BB#C_PRO238,81BC#C_PRO238
	LOG "Sysid checking use CDT Rule"
	if InStr(1,SysIdList,SysId&"#"&ChassisFB,vbTextCompare) <> 0 then SysIDChk = 1
Else
	REM (CNB) 810A,810B,810C,810D,79B1,79B2,79B3
	LOG "Sysid checking use CNB Rule"
	if InStr(1,SysIdList,SysId,vbTextCompare) <> 0               then SysIDChk = 1
End If
LOG ">>>>>>> SysIDChk=" & SysIDChk



rem ----- Summary SysLock ----------------------------------------------------------------------
Dim SysLock : SysLock = "FALSE"
If SysIDChk=1 AND OSFlagChk = 1 AND OptionCodeChk = 1 AND ExtraChk = 1 AND EditionServicingChk = 1 Then
'If SysIDChk=1 AND OSFlagChk = 1 AND OptionCodeChk = 1 AND ExtraChk = 1 Then
	SysLock = "TRUE"
End If
LOG4SSRM "SystemLock=" & SysLock
LOG      "SystemLock=" & SysLock
if SysLock then
	LOG "Result=PASSED"
else
	LOG "Result=FAILED"
end if

LOG4SSRM "---Print for SSRM parsing std output---------------------"
LOG4SSRM "ImageSku=" & OSSku_IMG
LOG4SSRM "UnitSku=" & OSSku_UUT
LOG4SSRM "OSSupport=" & OSFlagChk

LOG4SSRM "SystemId=" & SysId
LOG4SSRM "ChassisFB=" & ChassisFB
LOG4SSRM "SystemIdList=" & SysIdList
LOG4SSRM "SysIDChk=" & SysIDChk

LOG4SSRM "ImageOptionCode=" & OptionCodeS_IMG
LOG4SSRM "UnitOptionCode="  & OptionCodeS_UUT
LOG4SSRM "OptionCodeCheck=" & OptionCodeChk

LOG4SSRM "ImageCompact=" & isCompact_IMG
LOG4SSRM "UnitCompact="  & isCompact_UUT
LOG4SSRM "ImageCloudOS=" & isCloudOS_IMG
LOG4SSRM "UnitCloudOS="  & isCloudOS_UUT
LOG4SSRM "ExtraFeatureCheck=" & ExtraChk
LOG4SSRM "----------------------------------------------------------"

LOG vbCrLf & vbCrLf & vbCrLf

WScript.Quit (0)















rem ====== SUBs ====================================================================


Function Trimend(word, trimchar)
	Dim newword : newword = word
	if Right(newword,1) = trimchar then
		newword = Left(newword, Len(word) - Len(trimchar))
		newword = Trimend(newword, trimchar)
	end if 
	Trimend = newword
End Function

Function WriteFile(strOut,file)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Dim objFs : Set objFs = CreateObject("Scripting.FileSystemObject")
	Dim objFile : Set objFile = objFs.OpenTextFile(file, ForAppending, True)
	If IsObject(objFile) <> True Then Exit Function
	objFile.Write strOut & VbCrLf
	objFile.Close
	Set objFile = Nothing
End Function

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

Function ParseFeatureByte(FB, iniPath, sectionName, prefix)
	If StrComp(Left(FB,1),".") = 0 Then
		ParseFeatureByte = "NA"
		Exit Function
	End If
	Dim optionValue : OptionValue = ReadIniValue_CaseSensitive(iniPath, sectionName, Left(FB, 2), "0")
	LOG Left(FB, 2) & "=" & OptionValue
	If StrComp(Left(OptionValue, 2), prefix, vbTextCompare) = 0 Then
		ParseFeatureByte = OptionValue
	Else
		ParseFeatureByte = ParseFeatureByte(Right(FB, Len(FB)-2), iniPath, sectionName, prefix)
	End If
End Function


Function FindDmiIniPath(dmiRootFldrPath)
	Dim objDmiSubFldr
	Dim latestDate : latestDate = 0
	For Each objDmiSubFldr in fso.GetFolder(dmiRootFldrPath).SubFolders
		REM WScript.Echo "Find [" & objDmiSubFldr.Name & "] in " & dmiRootFldrPath
		Dim possibleDmiIniPath : possibleDmiIniPath = objDmiSubFldr.Path & "\DMI.INI"
		Dim currentFldrNameDate : currentFldrNameDate = CLng(objDmiSubFldr.Name)
		If fso.FileExists(possibleDmiIniPath) AND currentFldrNameDate > latestDate Then
			FindDmiIniPath = possibleDmiIniPath
			latestDate = currentFldrNameDate
		End If
	Next
	REM WScript.Echo "FindDmiIniPath=" & FindDmiIniPath
End Function

Function Help()
	WScript.Echo vbCrLf & "HP CNB Image Information Collection, Version 1.00,A2" & vbCrLf & "Copyright (c) 2010 Hewlett-Packard - All Rights Reserved" & vbCrLf
	WScript.Echo "Syntax: CScript.exe /nologo getimginfo.vbs" & vbCrLf
	WScript.Echo "Ex: CScript.exe /nologo getimginfo.vbs"
	WSCript.Quit(1)
End Function

Function ReadIniValue_CaseSensitive(inifile,section,key,default)
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
		 REM If StrComp(Left(Line, Len(key)), key, vbTextCompare) = 0 Then
		 If StrComp(Left(Line, Len(key)), key, vbBinaryCompare) = 0 Then
			strValue = Mid(Line, Len(key)+1)
		 End If
	  Next
   End If
   ReadIniValue_CaseSensitive = strValue
End Function



Sub MakeDirByFilePath(path)
	Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
	Dim idxBackSlash : idxBackSlash = InStrRev(path, "\")
	if idxBackSlash > 0 then MkDirRecurse Left(path, idxBackSlash-1), fso
	set fso = Nothing
End Sub


Sub MkDirRecurse(path, fso)
	'WSCript.echo "MkDirRecurse(" & path & ",fso)"
	Dim idxSlash : idxSlash = InStrRev(path, "\")
	if idxSlash > 0 then MkDirRecurse Left(path, idxSlash-1), fso
	if not fso.FolderExists(path) then fso.CreateFolder path
End Sub


Sub Err(msg)
	WScript.echo msg
	'WriteFile "[" & FormatDateTime(Time, 3) & "] " & msg, LogFile
End Sub

Sub LOG4SSRM(msg)
	WScript.echo msg
	LOG msg
End Sub

Sub LOG(msg)
	'WScript.echo msg
	WriteFile "[" & FormatDateTime(Time, 3) & "] " & msg, LogFile
End Sub

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


Function FB_Search(FB, token)
	If StrComp(Left(FB,1),".") = 0 Then
		FB_Search = false
		Exit Function
	End If
	Dim othersFB : othersFB = Right(FB, Len(FB)-2)
	Dim singleFB : singleFB = Left(FB, 2)
	'WScript.Echo "othersFB=" & othersFB
	'WScript.Echo "singleFB=" & singleFB
	if StrComp(singleFB, token) = 0 then
		FB_Search = true
		'WScript.Echo "FOUND"
		Exit Function
	Else
		FB_Search = FB_Search(othersFB, token)
	End If
End Function

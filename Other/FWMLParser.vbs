Option Explicit
On Error Resume Next

const UTIL_VERSION = "2.00,A,3"
REM =========================================================
REM 

REM Version: v2.00,A,3
REM Date: 2018/10/31
REM Author: tony.wu@hp.com
REM Update items: 1. Skip OA3 if no OA3 feature byte
REM               2. Clear patch PN if Linux/FreeDOS
REM
REM Version: v2.00,A,2
REM Date: 2018/10/31
REM Author: tony.wu@hp.com
REM Update items: 1. For OA3, check both C:\System.sav\WDT and C:\System.sav\Exclude\PINPROC\WDT
REM               2. Fix count issue if no FWML in SWPO
REM
REM Version: v2.00,A,1
REM Date: 2018/10/29
REM Author: tony.wu@hp.com
REM Update items: 1. Directly/Immediately trigger OA3
REM               2. Add FWMLParser version to PRA.ini
REM               3. Add OA3=Y/N to PRA.ini
REM               4. Remove file path arguments: /PRA and /FWML, force to c:\HP\Support\PRA.INI and c:\%FWML%
REM               5. Remove arguments: /LOG and /OVERWRITE (Default log name changed)
REM 
REM Version: v1.00,A,7
REM Date: 2018/7/9
REM Author: tony.wu@hp.com
REM Update items: 1. Align DPS Twekas to use Win32_PnPEntity
REM               2. If condition is with '/' and not SYSID\, treat as PNPID\, this is to support any new coming PNPID type
REM
REM Version: v1.00,A,6
REM Date: 2018/7/7
REM Author: tony.wu@hp.com
REM Update items: 1. Fix PNPID filter Win32_SignedPNPDriver not support in WinPE.
REM               2. Extend ID_PNPID
REM               3. Extend support condition / blk count from 999 to 99999
REM
REM Version: v1.00,A,5
REM Date: 2018/6/22
REM Author: tony.wu@hp.com
REM Update items: 1. Support PNPID filter
REM Update items: 2. Change default log path from c:\system.sav\logs\PRA\ to c:\system.sav\logs\ 
REM
REM Version: v1.00,A,4
REM Date: 2018/3/1
REM Author: tony.wu@hp.com
REM Update items: 1. Only keep factory updates which sortorder is 2
REM               2. Also exclude by name: "Factory Update 3PP File Base Hardware Hash Report", it should control by /SkipOA3
REM               3. Default log path change to <UP>\System.sav\log\PRA\ or <%~d0>\System.sav\log\PRA\
REM
REM Version: v1.00,A,3
REM Date: 2018/2/28
REM Author: tony.wu@hp.com
REM Update items: 1. Check region with SRC of BID
REM               2. Add SWPO file name (arg input)
REM               3. The PO/image do not support FWML, there's no feature byte: DL_FW
REM               4. "Output" rename to PRA
REM               5. Filter SSID and region, the script must be exected in target machine
REM               6. Support AND rule in a BLK condition, ex: PN.AB2,CT.AIO
REM               7. Supplemental Disc Add will be default filtered out by name.
REM               8. FWML path must be figured out in args.
REM
REM Version: v1.00,A,2
REM Date: 2018/1/30
REM Author: tony.wu@hp.com
REM Update items: 1. Arg figure out output folder: FactoryUpdate.LST and FlatFile.LST
REM               2. Input PO file, fwml file should be at the same folder
REM               3. Feature Byte read from PO
REM               4. Fill in FactoryUpdate 3PP Hardware Hash Report to list instead of directly modifying the UBRER.INI
REM               5. Write into FWML file name into output
REM
REM Version: v1.00,A,1
REM Date: 2018/1/22
REM Author: tony.wu@hp.com
REM Update items: 1. Initial release
REM
REM =========================================================
















REM =========================================================
const PATH_PRA_INI = "HP\Support\PRA.ini"
const PATH_WDT_1 = "System.sav\WDT"
const PATH_WDT_2 = "System.sav\Exclude\PINPROC\WDT"
const SHAREDATA = "sharedata.ini"
const UBERINI =   "uber.ini"
'const OA3_FU_PN = "P00PPD-B2D"
const EXCLUDE_FUNAME_LIST = "Supplemental Disc Addon;Factory Update 3PP File Base Hardware Hash Report"
const FULIST_SECTION_NAME= "FU_PN_LIST"
const FULIST_ITEM_PREFIX = "FACTORY_UPDATE_PN_"
const ID_REGION = "PN."
const ID_SYSID  = "SYSID\"
const FB_FDOS  = "6E"
const FB_LINUX = "6D"
const FB_OA3   = "aq"
REM =========================================================

const ITEM_MAX_COUNT=99999
const LINE_MAX_LENGTH=300
const MYLONGTAB = "                                "
Const ForReading = 1, ForWriting = 2, ForAppending = 8


Dim objArgs : Set objArgs = Wscript.Arguments
If objArgs.Count < 1 OR objArgs.Count > 2 Then
   Help()
   WSCript.Quit 1
End If

Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

Dim sLocale : sLocale = "en-us"
sLocale = getLocale()
'LOG "Current Locale: " & sLocale
SetLocale("en-us")
'LOG "Current Locale: " & sLocale

Dim UP : UP = fso.GetDriveName(Wscript.ScriptFullName)

Dim bFWMLinPO  : bFWMLinPO  = false
Dim bSkipOA3   : bSkipOA3   = false
Dim bClnPnList : bClnPnList = false
Dim PO_path    : PO_path    = ""
Dim FWML_path  : FWML_path  = ""
Dim FWML_name  : FWML_name  = ""
DIM PRA_path   : PRA_path = UP & "\" & PATH_PRA_INI
Dim pathSharedataIni : pathSharedataIni = "Unknown"
Dim pathUberIni      : pathUberIni      = "Unknown"

Dim arg
for each arg in objArgs
	If InStr(1, arg, "/PO=", vbTextCompare) >0 then
		PO_path = Mid(arg, InStr(arg, "=")+1)
	elseif InStr(1, arg, "/SKIPOA3", vbTextCompare) >0 then
		bSkipOA3 = true
	end if
next

Dim LogFile : LogFile = UP & "\System.sav\logs\PRA_" &  TimeStamp() & ".log"
WScript.Echo "LogFile="""& LogFile &""""


rem         'TESTING_SUB_FUNCTION
rem         WScript.Echo "ACPI\ELAN0732\4&329070F&0"
rem         WScript.Echo getDeviceIDList()
rem         If InStr(1, getDeviceIDList(), "ACPI\ELAN0732\4&329070F&0", vbTextCompare)>0  then 
rem         	WScript.Echo "Match"
rem         else
rem         	WScript.Echo "Mismatch"
rem         End If
rem         Wscript.Quit 1


'Check args
if (Len(PO_path)<=0) or (not fso.FileExists(PO_path)) then
	WScript.Echo "ERROR: PO path not found."
	Help()
	SetLocale(sLocale)
	WScript.Quit(1)
end if
if fso.FileExists(PRA_path) then
	WScript.Echo "ERROR: " & PRA_path & " already exists."
	Help()
	SetLocale(sLocale)
	WScript.Quit(1)
end if









LOG "Utility Version=" & UTIL_VERSION
LOG "PO_path=[" & PO_path & "]"
PO_path=fso.GetFile(PO_path).Path 'Get full path
LOG "PO_path=[" & PO_path & "]"
LOG "PRA_path=[" & PRA_path & "]"
MakeDirByFilePath PRA_path


DIM strFB : strFB = ReadIniValue_woSection(PO_path,";Feature_Byte","")
if Len(strFB)<=0 then
	WScript.Echo "ERROR: Feature Byte is not found in PO"
	SetLocale(sLocale)
	WScript.Quit(1)
end if
LOG "Feature Byte=[" & strFB & "]"

'Check non OA3 OS
if not FB_Search(strFB, FB_OA3) then 
	LOG "No OA3 feature byte, set SkipOA3 as true"
	bSkipOA3 = true
end if

'Check Linux FreeDOS feature byte
if FB_Search(strFB, FB_FDOS) then
	LOG "With FreeDOS feature byte, clear patch PN list"
	bClnPnList = true
end if
if FB_Search(strFB, FB_LINUX) then
	LOG "With Linux feature byte, clear patch PN list"
	bClnPnList = true
end if


FWML_path = "NULL"
FWML_name = ReadIniValue_woSection(PO_path,";FWML","NULL")
LOG "FWML_name=[" & FWML_name & "]"
if StrComp(FWML_name, "NULL", vbTextCompare)=0 then
	WScript.Echo "No FWML required from SWPO"
	bFWMLinPO = false
else
	bFWMLinPO = true
	if fso.FileExists(UP & "\"    & FWML_name & ".ini") then FWML_path = UP & "\"    & FWML_name & ".ini"
	if fso.FileExists(UP & "\CNB" & FWML_name & ".ini") then FWML_path = UP & "\CNB" & FWML_name & ".ini"
	if fso.FileExists(UP & "\CDT" & FWML_name & ".ini") then FWML_path = UP & "\CDT" & FWML_name & ".ini"
	if StrComp(FWML_path, "NULL", vbTextCompare)=0 then
		WScript.Echo "ERROR: FWML name is not found in PO or not align to FWML path provided by argument"
		SetLocale(sLocale)
		WScript.Quit(1)
	end if
end if
LOG "FWML_path=[" & FWML_path & "]"

if NOT bSkipOA3 then
	Dim tmpPathWdtFldr : tmpPathWdtFldr = ""
	if fso.FolderExists(UP & "\" & PATH_WDT_1) then tmpPathWdtFldr = UP & "\" & PATH_WDT_1
	if fso.FolderExists(UP & "\" & PATH_WDT_2) then tmpPathWdtFldr = UP & "\" & PATH_WDT_2
	if Len(tmpPathWdtFldr) <=0 then
		WScript.Echo "ERROR: WDT folder not exists."
		Help()
		SetLocale(sLocale)
		WScript.Quit(1)
	end if
	LOG "tmpPathWdtFldr=[" & tmpPathWdtFldr & "]"

	pathSharedataIni = tmpPathWdtFldr & "\" & SHAREDATA
	pathUberIni      = tmpPathWdtFldr & "\" & UBERINI
	LOG "pathSharedataIni=[" & pathSharedataIni & "]"
	LOG "pathUberIni=[" & pathUberIni & "]"
	if not fso.FileExists(pathSharedataIni) then
		WScript.Echo "ERROR: " & pathSharedataIni & " not exists."
		Help()
		SetLocale(sLocale)
		WScript.Quit(1)
	end if
	if not fso.FileExists(pathUberIni) then
		WScript.Echo "ERROR: " & pathUberIni & " not exists."
		Help()
		SetLocale(sLocale)
		WScript.Quit(1)
	end if
	if not PRAOA3() then
		WScript.Echo "ERROR: Failed to trigger OA3."
		Help()
		SetLocale(sLocale)
		WScript.Quit(1)
	end if
end if




DIM strBID : strBID = ReadIniValue_woSection(PO_path,";Build_ID","")
if Len(strBID)<=0 then
	WScript.Echo "ERROR: BuildID is not found in PO"
	SetLocale(sLocale)
	WScript.Quit(1)
end if
LOG "BuildID=[" & strBID & "]"

Dim strRegion : strRegion = Mid(strBID, InStr(1, strBID, "#S")+2, 3)
LOG "Region=[" & strRegion & "]"

Dim strSysid : strSysid = GetSSID()
LOG "strSysid=[" & strSysid & "]"

Dim strPnpIdLst : strPnpIdLst = getDeviceIDList()
LOG "strPnpIdLst=[" & strPnpIdLst & "]"


LOG ""
LOG "--- Basic Info -------------"
Dim strOA3 : strOA3 = "N" 
if not bSkipOA3 then strOA3 = "Y"
WriteFile "[INFO]",PRA_path
WriteFile "VERSION="   & UTIL_VERSION,             PRA_path
WriteFile "BID="       & strBID,                   PRA_path
WriteFile "FB="        & strFB,                    PRA_path
WriteFile "REGION="    & strRegion,                PRA_path
WriteFile "SYSID="     & strSysid,                 PRA_path
WriteFile "PO="        & fso.GetBaseName(PO_path), PRA_path
WriteFile "FWML="      & FWML_name,                PRA_path
WriteFile "PO_path="   & PO_path,                  PRA_path
WriteFile "FWML_path=" & FWML_path,                PRA_path
WriteFile "OA3="       & strOA3,                   PRA_path
'WriteFile "PNPID_LST=" & strPnpIdLst,PRA_path
WriteFile "",PRA_path


WriteFile "[" & FULIST_SECTION_NAME & "]",PRA_path
if not bFWMLinPO then
	WriteFile "TotalCount=0",PRA_path
elseif bClnPnList then
	WriteFile "TotalCount=0",PRA_path
Else
	Dim FULIST_idx : FULIST_idx = 0
	Dim FU_PN, FU_NAME, FU_SEQ, FU_isINCLUDED
	Dim cndtn_result_detail
	Dim cndtn, cndtn_2C
	
	'bXXXXX boolean indicator
	' true/false (true: match, false: mismatch)
	'iXXXXX int indicator
	' -1: No condition defined
	'  0: Condition defined and match
	'  1: Condition defined but mismatch
	Dim bMatch_SEQ
	Dim bMatch_Name
	Dim iMatch_Cndtn
	
	Dim i, j
	For i = 1 to ITEM_MAX_COUNT
		FU_PN = ReadIniValue_index(FWML_path,"FILESETS","Blk_", i)
		if StrComp(FU_PN, "SCRIPIT_INFO:INI_ILLEGAL_FORMAT", vbTextCompare) = 0 then
			LOG "ERROR: FILESETS has no BLK_" & i & " but there's BLK_" & (i+1) & ", please double check FWML."
			SetLocale(sLocale)
			Wscript.Quit(1)
		ElseIf StrComp(FU_PN, "SCRIPIT_INFO:END", vbTextCompare) = 0 then
			'LOG "FILESETS end at BLK_" & i
			Exit For
		Else
			FU_isINCLUDED = false
	
			LOG ""
			LOG ""
			LOG ""
			LOG "--- BLK_" & i & " -------------"
			LOG "PN=" & FU_PN
	
			'=== Check Not Supplemental Disc Addon ======================================================================================
			bMatch_Name = true
			FU_NAME=ReadIniValue(FWML_path,"COMPONENTS",      "Blk_" & i ,"NOT_FOUND")
			LOG "Name=" & FU_NAME
			if isExcludedName(FU_NAME, EXCLUDE_FUNAME_LIST) then
				LOG "Identified as a exclude name, never included in list"
				bMatch_Name = false
			End If
			LOG ">>>>>>> Not Exclude Name match=" & bMatch_Name
			
			'=== Check Sort Order ================================================================================
			bMatch_SEQ = false
			FU_SEQ= CInt(ReadIniValue(FWML_path,"INSTALL_SEQUENCE","Blk_" & i ,"NOT_FOUND"))
			LOG "Sort Order=" & FU_SEQ
			if FU_SEQ=2 then bMatch_SEQ = true
			LOG ">>>>>>> Sort Order match=" & bMatch_SEQ
			
			'=== Check Condition: OR ================================================================================
			iMatch_Cndtn = -1
			cndtn_result_detail = ""
			For j = 1 to ITEM_MAX_COUNT
				cndtn = ReadIniValue_index(FWML_path, "CONDITIONS", "BLK_" & i & "_", j)
				if StrComp(cndtn, "SCRIPIT_INFO:INI_ILLEGAL_FORMAT", vbTextCompare) = 0 then
					LOG "ERROR: CONDITIONS has no BLK_" & i & "_" & j & " but there's BLK_" & i & "_" & j & ", please double check FWML."
					SetLocale(sLocale)
					Wscript.Quit(1)
				ElseIf StrComp(cndtn, "SCRIPIT_INFO:END", vbTextCompare) = 0 then
					'LOG "CONDITIONS of BLK_" & i & "  end at BLK_" & i & "_" & j
					Exit For
				Else
					LOG "Condition(AND rule if there's multiple sub-condition)=" & cndtn
					if iMatch_Cndtn < 0 then iMatch_Cndtn = 1
					'=== Check Condition: AND ================================================================================
					if ChkSingleContition(cndtn) then iMatch_Cndtn = 0
				End If
			Next
			if Len(cndtn_result_detail)>LINE_MAX_LENGTH then cndtn_result_detail = vbCrLf & MYLONGTAB & Replace(cndtn_result_detail, " || ", " || " & vbCrLf & MYLONGTAB ) & vbCrLf & MYLONGTAB
			LOG "ConditionsDetails={" & cndtn_result_detail & "Result=" & iMatch_Cndtn & "}"
			LOG ">>>>>>> Condition Match=" & iMatch_Cndtn
			
			if iMatch_Cndtn<=0 and bMatch_SEQ and bMatch_Name then FU_isINCLUDED = true
			LOG ""
			LOG ">>>>>>>>>>>>> To be include into list=" & FU_isINCLUDED
			if FU_isINCLUDED then
				FULIST_idx = FULIST_idx + 1
				WriteFile FULIST_ITEM_PREFIX & FULIST_idx & "=" & FU_PN,PRA_path
			end if
		End If
	Next
	WriteFile "TotalCount=" & FULIST_idx,PRA_path
end if

LOG ""
LOG "------------------------------------------------"
LOG "Completed"

SetLocale(sLocale)
WScript.Quit(0)













Function ReadIniValue_index(inifile, section, key_prefix, idx)
	ReadIniValue_index = ReadIniValue(inifile,section,key_prefix & idx, "NOT_FOUND")
	'WScript.Echo vbTab & vbTab & "ReadIniValue_index=[" & ReadIniValue_index & "]"
	if StrComp(ReadIniValue_index, "NOT_FOUND", vbTextCompare) = 0 then
		if StrComp(ReadIniValue(inifile,section,key_prefix & (idx+1) ,"NOT_FOUND"), "NOT_FOUND", vbTextCompare) <> 0 then
			ReadIniValue_index = "SCRIPIT_INFO:INI_ILLEGAL_FORMAT"
		Else
			ReadIniValue_index = "SCRIPIT_INFO:END"
		End If
	End If
End Function


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

Function ReadIniKey(inifile,section,value,default)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Dim fso, objFile, strText, strSection, strKey, PosSection, PosEndSection, PosValue, PosEndValue, Line
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFile = fso.OpenTextFile(inifile, ForReading, False, TristateUseDefault)
	strText = objFile.ReadAll
	objFile.Close
	set objFile = Nothing

	Dim tmpValue : tmpValue = "=" & value
	strKey = default
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
		For Each Line In strSection
			If StrComp(Right(Line, Len(tmpValue)), tmpValue, vbTextCompare) = 0 Then
				strKey = Left(Line, Len(Line)-Len(tmpValue))
			End If
		Next
	End If
	ReadIniKey = strKey
End Function

Function ReadIniValue_woSection(inifile,key,default)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Dim fso, objFile, strText, strValue, Line
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFile = fso.OpenTextFile(inifile, ForReading, False, TristateUseDefault)
	strText = objFile.ReadAll
	objFile.Close
	set objFile = Nothing
	
	strValue = default
	key = key & "="
	For Each Line In split(strText, vbCrLf)
		If StrComp(Left(Line, Len(key)), key, vbTextCompare) = 0 Then
			strValue = Trim(Mid(Line, Len(key)+1))
		End If
	Next
	ReadIniValue_woSection = strValue
End Function

Function ReadAllContent(file,default)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Dim fso, objFile, strText, strSection, strValue, PosSection, PosEndSection, PosValue, PosEndValue, Line
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFile = fso.OpenTextFile(inifile, ForReading, False, TristateUseDefault)
	strText = objFile.ReadAll
	objFile.Close
	set objFile = Nothing
	ReadAllContent = strText
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


Function FBSearch(FB, FBToken)
	If StrComp(Left(FB,1),".") = 0 Then
		FBSearch = false
		Exit Function
	End If
	Dim othersFB : othersFB = Right(FB, Len(FB)-2)
	If StrComp(Left(FB, 2), FBToken, vbBinaryCompare) = 0 Then
		FBSearch = true
	Else
		FBSearch = FBSearch(othersFB, FBToken)
	End If
End Function



Function TimeStamp()
	Dim today : today = DatePart("yyyy", (date)) & num_wLZ(DatePart("m", (date)),2) & num_wLZ(DatePart("d", (date)),2)
	Dim Time_now : Time_now = num_wLZ(Hour(time),2) & num_wLZ(Minute(time),2) & num_wLZ(Second(time),2)
	TimeStamp = today & "_" & Time_now
End Function

Function num_wLZ(number, expectLength)
	num_wLZ = number
	while Len(num_wLZ) < expectLength 
		num_wLZ = "0" & num_wLZ
	wend
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

Function GetSSID()
	GetSSID = ""
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Dim colItems : Set colItems = objWMIService.ExecQuery("SELECT Product FROM Win32_BASEBOARD")
	Dim objItem
	For Each objItem In colItems
		GetSSID = objItem.Product
	Next
	set objWMIService = nothing
	set colItems = nothing
End Function




Function isRightCharsMatch(strEntire, strToBeCompare)
	isRightCharsMatch = false
	if StrComp(Right(strEntire, Len(strToBeCompare)), strToBeCompare, vbTextCompare)=0 then isRightCharsMatch = true
End Function

Function isLeftCharsMatch(strEntire, strToBeCompare)
	'WScript.Echo "isLeftCharsMatch(" & strEntire & ", " & strToBeCompare & ")"
	isLeftCharsMatch = false
	if StrComp(Left(strEntire, Len(strToBeCompare)), strToBeCompare, vbTextCompare)=0 then 
		isLeftCharsMatch = true
	End If
End Function


Function AppendString(str_prior, str_new, delims)
	If Len(str_prior) > 0 then str_new = " " & delims & " " & str_new
	AppendString = str_prior & str_new
End Function
Function AppendStringEx(str_prior, str_abbr, str_key, str_match, delims)
	AppendStringEx = AppendString(str_prior, "(" & str_abbr & ":" & str_key & "=" & str_match & ")", delims)
End Function


Function ChkSingleContition(sSingleCondition)
	ChkSingleContition = true
	Dim bIsMatch
	Dim subCondition
	Dim tmp_ResultDetail : tmp_ResultDetail = ""
	for each subCondition in Split(sSingleCondition, ",")
		subCondition = Trim(subCondition)
		if Len(subCondition)>0 then
			cndtn_2C=ReadIniKey(FWML_path,"Option_Map", subCondition ,"")
			If Len(cndtn_2C)<>2 then
				'LOG "Speical Condition"
				
				if isLeftCharsMatch(subCondition, ID_REGION) then 'Blk_2_14=PN.ABA,CT.AIO
					LOG "Special Condition: Region=""" & subCondition & """"
					bIsMatch = isRightCharsMatch(subCondition, strRegion)
					if Not bIsMatch then ChkSingleContition = false
					tmp_ResultDetail = AppendStringEx(tmp_ResultDetail, "REGION", subCondition, bIsMatch, "&&")
				elseif inStr(1, subCondition, "\", vbTextCompare) >0 then 'Blk_27_1=SYSID\4353 or Blk_27_1=PCI\VEN_8086&DEV_08B1
					if isLeftCharsMatch(subCondition, ID_SYSID) then 'Blk_27_1=SYSID\4353
						LOG "Special Condition: SysID=""" & subCondition & """"
						bIsMatch = isRightCharsMatch(subCondition, strSysid)
						if Not bIsMatch then ChkSingleContition = false
						tmp_ResultDetail = AppendStringEx(tmp_ResultDetail, "SYSID", subCondition, bIsMatch, "&&")
					else 'Blk_14_1=PCI\VEN_8086&DEV_08B1
						LOG "Special Condition: PNPID=""" & subCondition & """"
						bIsMatch = false
						if InStr(1, strPnpIdLst, subCondition, vbTextCompare) >0 then bIsMatch = true
						if Not bIsMatch then ChkSingleContition = false
						tmp_ResultDetail = AppendStringEx(tmp_ResultDetail, "PNPID", subCondition, bIsMatch, "&&")
					end if
				else
					LOG "ERROR: Undefined condition " & subCondition & " in Option_Map, please double check FWML."
					SetLocale(sLocale)
					Wscript.Quit(1)
				end if
			Else
				LOG "Common Condition: Feature Byte=""" & subCondition & """"
				bIsMatch = FBSearch(strFB, cndtn_2C)
				if not bIsMatch then ChkSingleContition = false
				tmp_ResultDetail = AppendStringEx(tmp_ResultDetail, subCondition, cndtn_2C, bIsMatch, "&&")
			End If
		end if
	Next
	cndtn_result_detail=AppendString(cndtn_result_detail, "[" & tmp_ResultDetail & "Result=" & ChkSingleContition & "]", "||")
End Function


Function isExcludedName(name, excludeNameList)
	'WScript.Echo "Checking " & name
	isExcludedName = false
	
	Dim excldName
	Dim idx, lastIdx
	Dim token, lastToken
	Dim isTheName
	
	for each excldName in Split(excludeNameList, ";")
		LOG "Compare to " & excldName
		isTheName = true
		idx=0
		lastIdx=0
		lastToken="N/A"
		for each token in Split(excldName, " ")
			idx = InStr(1, name, token, vbTextCompare)
			LOG "(" & idx & "," & lastIdx & ")=(""" & token & """,""" & lastToken & """)"
			if idx<=0 OR idx<=lastIdx then
				isTheName = false
				Exit For
			end if
			lastIdx=idx
			lastToken=token
		Next
		if isTheName then
			LOG """" & name & """ is a exclude name of """ & excldName & """"
			isExcludedName = true
			Exit Function
		else
			LOG """" & name & """ is NOT a exclude name of """ & excldName & """, check next"
		end if
	Next
End Function

Function getDeviceIDList()
	getDeviceIDList = ""
	Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	Dim colItems : Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_PnPEntity")
	Dim objItem
	For Each objItem In colItems
		If not IsNull(objItem) then
			'WScript.Echo "Checking getDeviceIDList[objItem=" & objItem.DeviceID & "]"
			If Len(getDeviceIDList) > 0 then getDeviceIDList = getDeviceIDList & vbCrLf
			getDeviceIDList = getDeviceIDList & Replace(UCase(objItem.DeviceID), "&AMP", "&", vbTextCompare)
		End If
	Next
	set objWMIService = nothing
	set colItems = nothing
End Function

Function isLeftCharsMatchAnyOf(strEntire, strLstToBeCompare)
	isLeftCharsMatchAnyOf = false
	Dim strToBeCompare, tmpStr
	'WScript.Echo "Checking strToBeCompare[strEntire=" & strEntire & "][strLstToBeCompare=" & strLstToBeCompare & "]"
	For Each strToBeCompare In Split(strLstToBeCompare, ";")
		tmpStr = Trim(strToBeCompare)
		'WScript.Echo "Checking strToBeCompare[tmpStr=" & tmpStr & "]"
		If Len(tmpStr)>0 then
			If isLeftCharsMatch(strEntire, tmpStr) then
				isLeftCharsMatchAnyOf = true
				Exit Function
			End If
		End If
	Next
End Function


'Clone from rwini.vbs
Function WriteIniValue(fileName, Section, KeyName, Value)
	Const ForReading = 1, ForWriting = 2, ForAppending = 8
	Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
	Dim fso, objFile, strText, strSection, strAfter, PosSection, PosEndSection
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objFile = fso.OpenTextFile(fileName, ForReading, True, TristateUseDefault)
	If objFile.AtEndOfStream Then
		strText = ""
	Else
		strText = objFile.ReadAll
	End If
	objFile.Close
	set objFile = Nothing  
	'WScript.Echo strText
	'Find section
	PosSection = InStr(1, strText, "[" & Section & "]", vbTextCompare)
	If PosSection>0 Then
		'Section exists. Find end of section
		PosEndSection = InStr(PosSection, strText, vbCrLf & "[")
		'?Is this last section?
		If PosEndSection = 0 Then PosEndSection = Len(strText)+1
		do while Mid(strText,PosEndSection-2,2) = vbCrLf
			PosEndSection=PosEndSection-2
		Loop

		'Separate section contents
		Dim OldSection, NewSection, Line
		Dim sKeyName
		Dim Found : Found = False
		OldSection = Mid(strText, PosSection, PosEndSection - PosSection)
		OldSection = split(OldSection, vbCrLf)

		'Temp variable To find a Key
		sKeyName = LCase(KeyName & "=")

		'Enumerate section lines
		For Each Line In OldSection
			If LCase(Left(Line, Len(sKeyName))) = sKeyName Then
				Line = KeyName & "=" & Value
				Found = True
			End If
			NewSection = NewSection & Line & vbCrLf
		Next

		If Found = False Then
			'key Not found - add it at the end of section
			
			NewSection = NewSection & KeyName & "=" & Value
		Else
			'remove last vbCrLf - the vbCrLf is at PosEndSection
			NewSection = Left(NewSection, Len(NewSection) - 2)
		End If

		'Combine pre-section, new section And post-section data.
		strText = Left(strText, PosSection-1) & NewSection & Mid(strText, PosEndSection)
	Else
		'Section Not found. Add section data at the end of file contents.
		If Right(strText, 2) <> vbCrLf And Len(strText)>0 Then 
			strText = strText & vbCrLf 
		End If
		strText = strText & "[" & Section & "]" & vbCrLf & _
		KeyName & "=" & Value
	End if
	Set objFile = fso.OpenTextFile(fileName, ForWriting, True, TristateUseDefault)
	
	objFile.Write strText
	objFile.Close
	set objFile = Nothing
End Function


Function PRAOA3()
	LOG "PRAOA3()"
	PRAOA3 = true
	'Cscript.exe /nologo "%~dp0rwinidq.vbs" write             "C:\System.sav\WDT\sharedata.ini" "OA3" "OA3PPReport" "Y"
	'Cscript.exe /nologo "%~dp0rwinidq.vbs" writeDoubleQuotes "C:\System.sav\WDT\uber.ini" "LABMODE" "OA3Check" ""TRUE""
	LOG "Set OA3PPReport=Y"
	WriteIniValue pathSharedataIni, "OA3",     "OA3PPReport", "Y"
	LOG "Set OA3Check=""TRUE"""
	WriteIniValue pathUberIni,      "LABMODE", "OA3Check",    """TRUE"""
	Dim val_OA3PPReport : val_OA3PPReport = ReadIniValue(pathSharedataIni, "OA3",     "OA3PPReport", "UNKNOW")
	Dim val_OA3Check    : val_OA3Check    = ReadIniValue(pathUberIni,      "LABMODE", "OA3Check",    "UNKNOW")
	LOG "val_OA3PPReport=[" & val_OA3PPReport & "]"
	LOG "val_OA3Check=[" & val_OA3Check & "]"
	If StrComp(val_OA3PPReport, "Y", vbTextCompare)<>0 then
		LOG "Failed to write [OA3] OA3PPReport=Y to " & pathSharedataIni
		PRAOA3 = false
	End If
	If StrComp(val_OA3Check, """TRUE""", vbTextCompare)<>0 then 
		LOG "Failed to write [LABMODE] OA3Check=""TRUE"" to " & pathUberIni
		PRAOA3 = false
	End If
	LOG "PRAOA3 is success=[" & PRAOA3 & "]"
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

Sub Help()
	WScript.Echo vbCrLf & vbCrLf & "==================================================" & vbCrLf
	WScript.echo "cscript.exe FWMLPaser.vbs /PO=<PO_PATH> [/SkipOA3]"
	WScript.echo "For OA3 enabling, ShareData.ini and Uber.ini must exists in image."
	WScript.Echo vbCrLf & "==================================================" & vbCrLf & vbCrLf
End Sub

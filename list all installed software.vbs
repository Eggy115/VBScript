' List All Installed Software



Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
strComputer = "."
strKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
strEntry1a = "DisplayName"
strEntry1b = "QuietDisplayName"
strEntry2 = "InstallDate"
strEntry3 = "VersionMajor"
strEntry4 = "VersionMinor"
strEntry5 = "EstimatedSize"

Set objReg = GetObject("winmgmts://" & strComputer & _
 "/root/default:StdRegProv")
objReg.EnumKey HKLM, strKey, arrSubkeys
WScript.Echo "Installed Applications" & VbCrLf
For Each strSubkey In arrSubkeys
  intRet1 = objReg.GetStringValue(HKLM, strKey & strSubkey, _
   strEntry1a, strValue1)
  If intRet1 <> 0 Then
    objReg.GetStringValue HKLM, strKey & strSubkey, _
     strEntry1b, strValue1
  End If
  If strValue1 <> "" Then
    WScript.Echo VbCrLf & "Display Name: " & strValue1
  End If
  objReg.GetStringValue HKLM, strKey & strSubkey, _
   strEntry2, strValue2
  If strValue2 <> "" Then
    WScript.Echo "Install Date: " & strValue2
  End If
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry3, intValue3
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry4, intValue4
  If intValue3 <> "" Then
     WScript.Echo "Version: " & intValue3 & "." & intValue4
  End If
  objReg.GetDWORDValue HKLM, strKey & strSubkey, _
   strEntry5, intValue5
  If intValue5 <> "" Then
    WScript.Echo "Estimated Size: " & Round(intValue5/1024, 3) & " megabytes"
  End If
Next

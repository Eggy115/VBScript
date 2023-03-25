' Windows Installer utility to manage installer policy settings
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the installer policy keys
' Policy can be configured by an administrator using the NT Group Policy Editor
'
Option Explicit

Dim policies(21, 4)
policies(1, 0)="LM" : policies(1, 1)="HKLM" : policies(1, 2)="Logging"              : policies(1, 3)="REG_SZ"    : policies(1, 4) = "Logging modes if not supplied by install, set of iwearucmpv"
policies(2, 0)="DO" : policies(2, 1)="HKLM" : policies(2, 2)="Debug"                : policies(2, 3)="REG_DWORD" : policies(2, 4) = "OutputDebugString: 1=debug output, 2=verbose debug output, 7=include command line"
policies(3, 0)="DI" : policies(3, 1)="HKLM" : policies(3, 2)="DisableMsi"           : policies(3, 3)="REG_DWORD" : policies(3, 4) = "1=Disable non-managed installs, 2=disable all installs"
policies(4, 0)="WT" : policies(4, 1)="HKLM" : policies(4, 2)="Timeout"              : policies(4, 3)="REG_DWORD" : policies(4, 4) = "Wait timeout in seconds in case of no activity"
policies(5, 0)="DB" : policies(5, 1)="HKLM" : policies(5, 2)="DisableBrowse"        : policies(5, 3)="REG_DWORD" : policies(5, 4) = "Disable user browsing of source locations if 1"
policies(6, 0)="AB" : policies(6, 1)="HKLM" : policies(6, 2)="AllowLockdownBrowse"  : policies(6, 3)="REG_DWORD" : policies(6, 4) = "Allow non-admin users to browse to new sources for managed applications if 1 - use is not recommended"
policies(7, 0)="AM" : policies(7, 1)="HKLM" : policies(7, 2)="AllowLockdownMedia"   : policies(7, 3)="REG_DWORD" : policies(7, 4) = "Allow non-admin users to browse to new media sources for managed applications if 1 - use is not recommended"
policies(8, 0)="AP" : policies(8, 1)="HKLM" : policies(8, 2)="AllowLockdownPatch"   : policies(8, 3)="REG_DWORD" : policies(8, 4) = "Allow non-admin users to apply small and minor update patches to managed applications if 1 - use is not recommended"
policies(9, 0)="DU" : policies(9, 1)="HKLM" : policies(9, 2)="DisableUserInstalls"  : policies(9, 3)="REG_DWORD" : policies(9, 4) = "Disable per-user installs if 1 - available on Windows Installer version 2.0 and later"
policies(10, 0)="DP" : policies(10, 1)="HKLM" : policies(10, 2)="DisablePatch"         : policies(10, 3)="REG_DWORD" : policies(10, 4) = "Disable patch application to all products if 1"
policies(11, 0)="UC" : policies(11, 1)="HKLM" : policies(11, 2)="EnableUserControl"    : policies(11, 3)="REG_DWORD" : policies(11, 4) = "All public properties sent to install service if 1"
policies(12, 0)="ER" : policies(12, 1)="HKLM" : policies(12, 2)="EnableAdminTSRemote"  : policies(12, 3)="REG_DWORD" : policies(12, 4) = "Allow admins to perform installs from terminal server client sessions if 1"
policies(13, 0)="LS" : policies(13, 1)="HKLM" : policies(13, 2)="LimitSystemRestoreCheckpointing" : policies(13, 3)="REG_DWORD" : policies(13, 4) = "Turn off creation of system restore check points on Windows XP if 1 - available on Windows Installer version 2.0 and later"
policies(14, 0)="SS" : policies(14, 1)="HKLM" : policies(14, 2)="SafeForScripting"     : policies(14, 3)="REG_DWORD" : policies(14, 4) = "Do not prompt when scripts within a webpage access Installer automation interface if 1 - use is not recommended"
policies(15, 0)="TP" : policies(15,1)="HKLM" : policies(15, 2)="TransformsSecure"     : policies(15, 3)="REG_DWORD" : policies(15, 4) = "Pin tranforms in secure location if 1 (only admin and system have write privileges to cache location)"
policies(16, 0)="EM" : policies(16, 1)="HKLM" : policies(16, 2)="AlwaysInstallElevated": policies(16, 3)="REG_DWORD" : policies(16, 4) = "System privileges if 1 and HKCU value also set - dangerous policy as non-admin users can install with elevated privileges if enabled"
policies(17, 0)="EU" : policies(17, 1)="HKCU" : policies(17, 2)="AlwaysInstallElevated": policies(17, 3)="REG_DWORD" : policies(17, 4) = "System privileges if 1 and HKLM value also set - dangerous policy as non-admin users can install with elevated privileges if enabled"
policies(18,0)="DR" : policies(18,1)="HKCU" : policies(18,2)="DisableRollback"      : policies(18,3)="REG_DWORD" : policies(18,4) = "Disable rollback if 1 - use is not recommended"
policies(19,0)="TS" : policies(19,1)="HKCU" : policies(19,2)="TransformsAtSource"   : policies(19,3)="REG_DWORD" : policies(19,4) = "Locate transforms at root of source image if 1"
policies(20,0)="SO" : policies(20,1)="HKCU" : policies(20,2)="SearchOrder"          : policies(20,3)="REG_SZ"    : policies(20,4) = "Search order of source types, set of n,m,u (default=nmu)"
policies(21,0)="DM" : policies(21,1)="HKCU" : policies(21,2)="DisableMedia"          : policies(21,3)="REG_DWORD"    : policies(21,4) = "Browsing to media sources is disabled"

Dim argCount:argCount = Wscript.Arguments.Count
Dim message, iPolicy, policyKey, policyValue, WshShell, policyCode
On Error Resume Next

' If no arguments supplied, then list all current policy settings
If argCount = 0 Then
	Set WshShell = WScript.CreateObject("WScript.Shell") : CheckError
	For iPolicy = 0 To UBound(policies)
		policyValue = ReadPolicyValue(iPolicy)
		If Not IsEmpty(policyValue) Then 'policy key present, else skip display
			If Not IsEmpty(message) Then message = message & vbLf
			message = message & policies(iPolicy,0) & ": " & policies(iPolicy,2) & "(" & policies(iPolicy,1) & ") = " & policyValue
		End If
	Next
	If IsEmpty(message) Then message = "No installer policies set"
	Wscript.Echo message
	Wscript.Quit 0
End If

' Check for ?, and show help message if found
policyCode = UCase(Wscript.Arguments(0))
If InStr(1, policyCode, "?", vbTextCompare) <> 0 Then
	message = "Windows Installer utility to manage installer policy settings" &_
		vbLf & " If no arguments are supplied, current policy settings in list will be reported" &_
		vbLf & " The 1st argument specifies the policy to set, using a code from the list below" &_
		vbLf & " The 2nd argument specifies the new policy setting, use """" to remove the policy" &_
		vbLf & " If the 2nd argument is not supplied, the current policy value will be reported"
	For iPolicy = 0 To UBound(policies)
		message = message & vbLf & policies(iPolicy,0) & ": " & policies(iPolicy,2) & "(" & policies(iPolicy,1) & ")  " & policies(iPolicy,4) & vbLf
	Next
	message = message & vblf & vblf & "Copyright (C) Microsoft Corporation.  All rights reserved."

	Wscript.Echo message
	Wscript.Quit 1
End If

' Policy code supplied, look up in array
For iPolicy = 0 To UBound(policies)
	If policies(iPolicy, 0) = policyCode Then Exit For
Next
If iPolicy > UBound(policies) Then Wscript.Echo "Unknown policy code: " & policyCode : Wscript.Quit 2
Set WshShell = WScript.CreateObject("WScript.Shell") : CheckError

' If no value supplied, then simply report current value
policyValue = ReadPolicyValue(iPolicy)
If IsEmpty(policyValue) Then policyValue = "Not present"
message = policies(iPolicy,0) & ": " & policies(iPolicy,2) & "(" & policies(iPolicy,1) & ") = " & policyValue
If argCount > 1 Then ' Value supplied, set policy
	policyValue = WritePolicyValue(iPolicy, Wscript.Arguments(1))
	If IsEmpty(policyValue) Then policyValue = "Not present"
	message = message & " --> " & policyValue
End If
Wscript.Echo message

Function ReadPolicyValue(iPolicy)
	On Error Resume Next
	Dim policyKey:policyKey = policies(iPolicy,1) & "\Software\Policies\Microsoft\Windows\Installer\" & policies(iPolicy,2)
	ReadPolicyValue = WshShell.RegRead(policyKey)
End Function

Function WritePolicyValue(iPolicy, policyValue)
	On Error Resume Next
	Dim policyKey:policyKey = policies(iPolicy,1) & "\Software\Policies\Microsoft\Windows\Installer\" & policies(iPolicy,2)
	If Len(policyValue) Then
		WshShell.RegWrite policyKey, policyValue, policies(iPolicy,3) : CheckError
		WritePolicyValue = policyValue
	ElseIf Not IsEmpty(ReadPolicyValue(iPolicy)) Then
		WshShell.RegDelete policyKey : CheckError
	End If
End Function

Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbLf & errRec.FormatText
	End If
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIIl9QYJKoZIhvcNAQcCoIIl5jCCJeICAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' /P2vh2Oob6dK4LJgJX9q0NCUL4Tr2otbVtqd2lYQZG6g
'' SIG '' gguBMIIFCTCCA/GgAwIBAgITMwAABHCfcxf1mw5kZgAA
'' SIG '' AAAEcDANBgkqhkiG9w0BAQsFADB+MQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
'' SIG '' aWduaW5nIFBDQSAyMDEwMB4XDTIyMDEyNzE5MzIyMVoX
'' SIG '' DTIzMDEyNjE5MzIyMVowfzELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEpMCcGA1UEAxMgTWljcm9zb2Z0IFdpbmRvd3MgS2l0
'' SIG '' cyBQdWJsaXNoZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IB
'' SIG '' DwAwggEKAoIBAQCVDWJkrimipfHoYzaKQr7se+gP72Rz
'' SIG '' FjWGY4S7MjjHKalJOqO4c5cJ18CyWs+VVT+bvCsyQ4Zn
'' SIG '' EK98JePG4DJ2Osxsyz9oY3Pc/E9XFa21fCrGJCiWLrCg
'' SIG '' hP3soGkAXyZDL3/nZr9CsFu1zThDnwEpjKIlMj/YA09/
'' SIG '' cNE25pmMnwpoWjZwaxXC1TM6SZHhOV271XlsgRn6OjFG
'' SIG '' Wdw7fmouu/tc6uu8MKKb+PZ0KAyDNa0gOjPUB9j6dZKD
'' SIG '' fr7p/D1CSvs2gSFN6f+v0XXUNp/sn1XH8x8nSSrt7LFJ
'' SIG '' Dq169hM9LaSGDjQ4iF71UH26lBr4pMJzF7jbLRNaMQr0
'' SIG '' XGqZAgMBAAGjggF9MIIBeTAfBgNVHSUEGDAWBgorBgEE
'' SIG '' AYI3CgMUBggrBgEFBQcDAzAdBgNVHQ4EFgQUFGZ2RTD4
'' SIG '' k22AqdFnqXZHPvpSWzQwVAYDVR0RBE0wS6RJMEcxLTAr
'' SIG '' BgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlv
'' SIG '' bnMgTGltaXRlZDEWMBQGA1UEBRMNMjI5OTAzKzQ2OTA2
'' SIG '' MjAfBgNVHSMEGDAWgBTm/F97uyIAWORyTrX0IXQjMubv
'' SIG '' rDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1p
'' SIG '' Y3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWND
'' SIG '' b2RTaWdQQ0FfMjAxMC0wNy0wNi5jcmwwWgYIKwYBBQUH
'' SIG '' AQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1p
'' SIG '' Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY0NvZFNpZ1BD
'' SIG '' QV8yMDEwLTA3LTA2LmNydDAMBgNVHRMBAf8EAjAAMA0G
'' SIG '' CSqGSIb3DQEBCwUAA4IBAQB6oT/eW1Dk6SENj/OuFcIf
'' SIG '' vLpUoLivcfVbQwNtpk+lNWF4cG3endkCyhiALEJMIPuS
'' SIG '' JG07OjcK0ZhKvvLKdrZ8EaYNwgsdBOQGhtEb9yFuG+X+
'' SIG '' O0VFSVq2o3yKJLImJh9WS6/BX13mdEuwASb8Zmtf613w
'' SIG '' C2sB7wdApagduw/5yoXEZsP5M0bFNFqTmt8xAVNyaNZK
'' SIG '' zjDyW8vfEQnASCraD+OSyKN6nPGUJSCDk6uGnWYLbsQ2
'' SIG '' uZARXw74kGACVNgnrtlL20vnbvbM6amV6nJs/MsvU2MV
'' SIG '' Xng/xM+J0GswE+XbxjO+bDVxwKIPLgGZ5t0ly0jsRjPj
'' SIG '' L4pVZl9AeScNMIIGcDCCBFigAwIBAgIKYQxSTAAAAAAA
'' SIG '' AzANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMx
'' SIG '' EzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1Jl
'' SIG '' ZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
'' SIG '' dGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2Vy
'' SIG '' dGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMTAwNzA2
'' SIG '' MjA0MDE3WhcNMjUwNzA2MjA1MDE3WjB+MQswCQYDVQQG
'' SIG '' EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UE
'' SIG '' BxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
'' SIG '' cnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29k
'' SIG '' ZSBTaWduaW5nIFBDQSAyMDEwMIIBIjANBgkqhkiG9w0B
'' SIG '' AQEFAAOCAQ8AMIIBCgKCAQEA6Q5kUHlntcTj/QkATJ6U
'' SIG '' rPdWaOpE2M/FWE+ppXZ8bUW60zmStKQe+fllguQX0o/9
'' SIG '' RJwI6GWTzixVhL99COMuK6hBKxi3oktuSUxrFQfe0dLC
'' SIG '' iR5xlM21f0u0rwjYzIjWaxeUOpPOJj/s5v40mFfVHV1J
'' SIG '' 9rIqLtWFu1k/+JC0K4N0yiuzO0bj8EZJwRdmVMkcvR3E
'' SIG '' VWJXcvhnuSUgNN5dpqWVXqsogM3Vsp7lA7Vj07IUyMHI
'' SIG '' iiYKWX8H7P8O7YASNUwSpr5SW/Wm2uCLC0h31oVH1RC5
'' SIG '' xuiq7otqLQVcYMa0KlucIxxfReMaFB5vN8sZM4BqiU2j
'' SIG '' amZjeJPVMM+VHwIDAQABo4IB4zCCAd8wEAYJKwYBBAGC
'' SIG '' NxUBBAMCAQAwHQYDVR0OBBYEFOb8X3u7IgBY5HJOtfQh
'' SIG '' dCMy5u+sMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBB
'' SIG '' MAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8G
'' SIG '' A1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYG
'' SIG '' A1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9z
'' SIG '' b2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0Nl
'' SIG '' ckF1dF8yMDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQRO
'' SIG '' MEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9z
'' SIG '' b2Z0LmNvbS9wa2kvY2VydHMvTWljUm9vQ2VyQXV0XzIw
'' SIG '' MTAtMDYtMjMuY3J0MIGdBgNVHSAEgZUwgZIwgY8GCSsG
'' SIG '' AQQBgjcuAzCBgTA9BggrBgEFBQcCARYxaHR0cDovL3d3
'' SIG '' dy5taWNyb3NvZnQuY29tL1BLSS9kb2NzL0NQUy9kZWZh
'' SIG '' dWx0Lmh0bTBABggrBgEFBQcCAjA0HjIgHQBMAGUAZwBh
'' SIG '' AGwAXwBQAG8AbABpAGMAeQBfAFMAdABhAHQAZQBtAGUA
'' SIG '' bgB0AC4gHTANBgkqhkiG9w0BAQsFAAOCAgEAGnTvV08p
'' SIG '' e8QWhXi4UNMi/AmdrIKX+DT/KiyXlRLl5L/Pv5PI4zSp
'' SIG '' 24G43B4AvtI1b6/lf3mVd+UC1PHr2M1OHhthosJaIxrw
'' SIG '' jKhiUUVnCOM/PB6T+DCFF8g5QKbXDrMhKeWloWmMIpPM
'' SIG '' dJjnoUdD8lOswA8waX/+0iUgbW9h098H1dlyACxphnY9
'' SIG '' UdumOUjJN2FtB91TGcun1mHCv+KDqw/ga5uV1n0oUbCJ
'' SIG '' SlGkmmzItx9KGg5pqdfcwX7RSXCqtq27ckdjF/qm1qKm
'' SIG '' huyoEESbY7ayaYkGx0aGehg/6MUdIdV7+QIjLcVBy78d
'' SIG '' TMgW77Gcf/wiS0mKbhXjpn92W9FTeZGFndXS2z1zNfM8
'' SIG '' rlSyUkdqwKoTldKOEdqZZ14yjPs3hdHcdYWch8ZaV4XC
'' SIG '' v90Nj4ybLeu07s8n07VeafqkFgQBpyRnc89NT7beBVaX
'' SIG '' evfpUk30dwVPhcbYC/GO7UIJ0Q124yNWeCImNr7KsYxu
'' SIG '' qh3khdpHM2KPpMmRM19xHkCvmGXJIuhCISWKHC1g2TeJ
'' SIG '' QYkqFg/XYTyUaGBS79ZHmaCAQO4VgXc+nOBTGBpQHTiV
'' SIG '' mx5mMxMnORd4hzbOTsNfsvU9R1O24OXbC2E9KteSLM43
'' SIG '' Wj5AQjGkHxAIwlacvyRdUQKdannSF9PawZSOB3slcUSr
'' SIG '' Bmrm1MbfI5qWdcUxghnMMIIZyAIBATCBlTB+MQswCQYD
'' SIG '' VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
'' SIG '' A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
'' SIG '' IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
'' SIG '' Q29kZSBTaWduaW5nIFBDQSAyMDEwAhMzAAAEcJ9zF/Wb
'' SIG '' DmRmAAAAAARwMA0GCWCGSAFlAwQCAQUAoIIBBDAZBgkq
'' SIG '' hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3
'' SIG '' AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQx
'' SIG '' IgQgG74cd1WwHhi9o902FBuqooOIxL5KiicynZhgW/cw
'' SIG '' OaswPAYKKwYBBAGCNwoDHDEuDCxQVi8xcDVsQkxIMEUy
'' SIG '' ZklwNVN1S3BkbDFnQ21LMEZHTTlHWDk5UUxJWHN3PTBa
'' SIG '' BgorBgEEAYI3AgEMMUwwSqAkgCIATQBpAGMAcgBvAHMA
'' SIG '' bwBmAHQAIABXAGkAbgBkAG8AdwBzoSKAIGh0dHA6Ly93
'' SIG '' d3cubWljcm9zb2Z0LmNvbS93aW5kb3dzMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBACh3N/4BtT/zSgOIFSVSIWi4Qsk3molf
'' SIG '' 5wBJxJ+Ahn6mIyeqDl6ofzO6nw1NkMZGsdxmthJqcaTS
'' SIG '' F5YkBO2APAcboQOwF6rNjL0oHlWoIQLwZ9iyMhP8c+lz
'' SIG '' vLcNgYrrhxcCSRZqHTJbZABs0GV8bqKkeb92r/d/9kRI
'' SIG '' BXXe4FM/eIxWgOZaqzkkynD4iWCGfqdQK6b0fZO043jT
'' SIG '' wYgu+OBxybzoNnPBZ5iswn4vbVM3SX2zX2Cu+xC0M6DG
'' SIG '' dUIzd9302A8W4PWe0pflJuano1urd+vbTVXBVVAMtgTF
'' SIG '' NFS8YzZDszW7xiBU4/TuokZQ4B6o2YbssiHYsMfDPjjJ
'' SIG '' Mvmhghb/MIIW+wYKKwYBBAGCNwMDATGCFuswghbnBgkq
'' SIG '' hkiG9w0BBwKgghbYMIIW1AIBAzEPMA0GCWCGSAFlAwQC
'' SIG '' AQUAMIIBUAYLKoZIhvcNAQkQAQSgggE/BIIBOzCCATcC
'' SIG '' AQEGCisGAQQBhFkKAwEwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' jYZfQWGAY7tp9cWmYtGK5pG0FOF5QWJ2bSwWH9tv5C8C
'' SIG '' BmLP9QXTSRgSMjAyMjA3MTYwODU2NTguODFaMASAAgH0
'' SIG '' oIHQpIHNMIHKMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
'' SIG '' V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
'' SIG '' A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYD
'' SIG '' VQQLExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25z
'' SIG '' MSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpERDhDLUUz
'' SIG '' MzctMkZBRTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUt
'' SIG '' U3RhbXAgU2VydmljZaCCEVcwggcMMIIE9KADAgECAhMz
'' SIG '' AAABnA+mTWHSnksoAAEAAAGcMA0GCSqGSIb3DQEBCwUA
'' SIG '' MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
'' SIG '' dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
'' SIG '' aWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
'' SIG '' Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTIx
'' SIG '' MTIwMjE5MDUxOVoXDTIzMDIyODE5MDUxOVowgcoxCzAJ
'' SIG '' BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
'' SIG '' DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
'' SIG '' ZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29m
'' SIG '' dCBBbWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRo
'' SIG '' YWxlcyBUU1MgRVNOOkREOEMtRTMzNy0yRkFFMSUwIwYD
'' SIG '' VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
'' SIG '' MIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA
'' SIG '' 21IqDBldSRY/rOtdrNNpirttyj1DbO9Tow3iRrcZExfa
'' SIG '' 0lk4rgPF4GJAAIv+fthX+wpOyXCkPR/1w9TisINf2x9x
'' SIG '' Najtc/F0ctD5aRoZsopYBOyrDr1vDyGQn9uNynZXYDq8
'' SIG '' ay/ByokKHTsErck+ZS1mKTLLk9nx/JPKIoY3uE5aVohT
'' SIG '' 2gii5xQ2gAdAnMuryHbR42AdSHt4jmT4rKri/rzXQse4
'' SIG '' DoQfIok5k3bFPDklKQvLQU3kyGD85oWsUGXeJqDZOqng
'' SIG '' icou34luH8l3R62d6LZoMcWuaV8+aVFK/nBI1fnMCGAT
'' SIG '' JGmOZBzPXOnRBpIB59GQyb3bf+eBTnUhutVsB4ePnr1I
'' SIG '' cL12geCwjGSHQreWnDnzb7Q41dwh8hTqeQFP6oAMBn7R
'' SIG '' 1PW67+BFMHLrXhACh+OjbnxNtJf1o5TVIe4AL7dsyjIz
'' SIG '' uM10cQlE4f6awUMFyYlGXhUqxF4jn5Lr0pQZ4sgGGGae
'' SIG '' ZDp2sXwinRmI76+ECwPd70CeqdjsdyB7znQj2gq/C7Cl
'' SIG '' XBacqfDBIYSUzPtS8KhyahQxeTtWfZo22L5t0fbz4ZBv
'' SIG '' kQyyqE6a+5k4JGk5Y3fcb5veDm6fAQ/R5OJj4udZrYC4
'' SIG '' rjfP+mmVRElWV7b0rjZA+Q5yCUHqyMuY2kSlv1tqwnvZ
'' SIG '' 4DQyWnUu0fehhkZeyCBN+5cCAwEAAaOCATYwggEyMB0G
'' SIG '' A1UdDgQWBBS7aQlnU12OXbXXZLKcvqMYwgP6sjAfBgNV
'' SIG '' HSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNV
'' SIG '' HR8EWDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29m
'' SIG '' dC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1l
'' SIG '' LVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYB
'' SIG '' BQUHAQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3
'' SIG '' Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jv
'' SIG '' c29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEp
'' SIG '' LmNydDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsG
'' SIG '' AQUFBwMIMA0GCSqGSIb3DQEBCwUAA4ICAQCnACqmIxhH
'' SIG '' M01jLPc9Ju2KNt7IKlRdy8iuoDjM+0whwCTfhb272ZEO
'' SIG '' d1ZL62VHdbBOmvU6BpXXCZzpgXOoroQZab3TdQSwUTvE
'' SIG '' Ekw9eN91U4+FwkHe9+8DQ9fnqihtwXY682w5LBMHxuL+
'' SIG '' ez4Kzf0+7Oz5BI1Bl3yIBUEJK/E0Ivvx2WfZEZTXHIHg
'' SIG '' AqpX2+Lhj8Z+bHYUD6MXTL5gt6hvQzjSeVLEvSrTvm3s
'' SIG '' vqIVEw2vS7xE6HOEM8uX7h49h9SbJgmihu/J16X1qcAS
'' SIG '' wcWWEqX5pdvaJzfI3Buyg/Jxkkv++jw5W9hjELL7/kWt
'' SIG '' CYC+hbRkRoGJhwqTOs1a3+Ff2vkqB3AvrXHRmJNmilOS
'' SIG '' jpb/nxRN59NuFfs+eLQwCkfc+/K3o3QgVqn78uXAVEPX
'' SIG '' Oft7pxw9PARKe6j9q4KaA/OerzQ4BMDu+5+xFk++p5fy
'' SIG '' Mq2ytpI2xy81DKYRaVyp1dX2FiSNvhP9Cx71xRhqheDr
'' SIG '' zAUcW6yVZ9N09g8uXW+rOU8yc0mkLwq12KgOByr7LUFp
'' SIG '' KpKbwR01/DNPfv78kW1Vzcaz3Xl8OqA9kOA5LMpAhX5/
'' SIG '' Ddo9i3YsRPcBuYopb+vXc7LxyDf4PQPfrYZAEAlW/Q1E
'' SIG '' jk2jCBoLDqg2BY4U+s3vZZIRxxr/xBCJMY/ZekuIalEM
'' SIG '' lnqxZGlFg13J2TCCB3EwggVZoAMCAQICEzMAAAAVxedr
'' SIG '' ngKbSZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJ
'' SIG '' BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
'' SIG '' DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
'' SIG '' ZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29m
'' SIG '' dCBSb290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEw
'' SIG '' MB4XDTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVow
'' SIG '' fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0G
'' SIG '' CSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57Ry
'' SIG '' IQt5osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg
'' SIG '' 4r25PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPP
'' SIG '' dzLAEBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFEx
'' SIG '' N6AKOG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTa
'' SIG '' mDu6GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyF
'' SIG '' Vk3v3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIck
'' SIG '' w+DJj361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50Zu
'' SIG '' yjLVwIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viS
'' SIG '' kR4dPf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37Zy
'' SIG '' L9t9X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0lBw0
'' SIG '' gg/wEPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8
'' SIG '' QmguEOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDw
'' SIG '' wvoSCtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2
'' SIG '' BjUYhEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y
'' SIG '' 7435UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJp
'' SIG '' xq57t7c+auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGC
'' SIG '' NxUBBAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTE
'' SIG '' mr6CkTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJd
'' SIG '' g/Tl0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGC
'' SIG '' N0yDfQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cu
'' SIG '' bWljcm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0
'' SIG '' b3J5Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkr
'' SIG '' BgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
'' SIG '' AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV
'' SIG '' 9lbLj+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEug
'' SIG '' SaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' L2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0w
'' SIG '' Ni0yMy5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUF
'' SIG '' BzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtp
'' SIG '' L2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNy
'' SIG '' dDANBgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwEx
'' SIG '' JFvhnnJL/Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/a
'' SIG '' ZGx80HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th54
'' SIG '' 2DYunKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZ
'' SIG '' GN5tggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2
'' SIG '' KFUn0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVC
'' SIG '' s/wMnosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8s
'' SIG '' CXgU6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0
'' SIG '' j/aRAfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2d
'' SIG '' Y3RILLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/
'' SIG '' HltEAY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nx
'' SIG '' t67I6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+CrvsQWY9
'' SIG '' af3LwUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcd
'' SIG '' FYmNcP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaM
'' SIG '' mdbhIurwJ0I9JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE
'' SIG '' 0ZDxyKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNY
'' SIG '' s6FwZvKhggLOMIICNwIBATCB+KGB0KSBzTCByjELMAkG
'' SIG '' A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAO
'' SIG '' BgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
'' SIG '' dCBDb3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0
'' SIG '' IEFtZXJpY2EgT3BlcmF0aW9uczEmMCQGA1UECxMdVGhh
'' SIG '' bGVzIFRTUyBFU046REQ4Qy1FMzM3LTJGQUUxJTAjBgNV
'' SIG '' BAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2Wi
'' SIG '' IwoBATAHBgUrDgMCGgMVAM3Zaerd8LP25xK25vXNDPvX
'' SIG '' b1NAoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNV
'' SIG '' BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
'' SIG '' HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEm
'' SIG '' MCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
'' SIG '' IDIwMTAwDQYJKoZIhvcNAQEFBQACBQDmfG3AMCIYDzIw
'' SIG '' MjIwNzE2MDY1MDQwWhgPMjAyMjA3MTcwNjUwNDBaMHcw
'' SIG '' PQYKKwYBBAGEWQoEATEvMC0wCgIFAOZ8bcACAQAwCgIB
'' SIG '' AAICFHYCAf8wBwIBAAICEacwCgIFAOZ9v0ACAQAwNgYK
'' SIG '' KwYBBAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgC
'' SIG '' AQACAwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0BAQUF
'' SIG '' AAOBgQBEp28wVPI2M+h0GcX2wgjSGPRhSll6S+y8nulR
'' SIG '' tIR8jhoBaHCrshTKMXOvwv8YktgdzjXSrvcRbvYkSOSE
'' SIG '' ZRmEGkUGxUDhdOhz2YiG5oXhIcd41vo8WEzgcJo7qlQj
'' SIG '' ZGYua2IkLg//mznwh0d+YqYGe+es3czRFzDN69ggA4Wa
'' SIG '' AjGCBA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMw
'' SIG '' EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
'' SIG '' b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
'' SIG '' b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1w
'' SIG '' IFBDQSAyMDEwAhMzAAABnA+mTWHSnksoAAEAAAGcMA0G
'' SIG '' CWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYL
'' SIG '' KoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIFOAW2GO
'' SIG '' eUDV65ba6LkdlcF1Y+lMD4CGLBPqdbpBMyjSMIH6Bgsq
'' SIG '' hkiG9w0BCRACLzGB6jCB5zCB5DCBvQQgNw9FhSCNLMo6
'' SIG '' EXf13hCBtFlCCs87suj+oTka29J6prwwgZgwgYCkfjB8
'' SIG '' MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
'' SIG '' bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWlj
'' SIG '' cm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNy
'' SIG '' b3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAZwP
'' SIG '' pk1h0p5LKAABAAABnDAiBCCAQylAR75IORWcPqETuXaH
'' SIG '' WBE0TWurg6Hv20zBuD0TxTANBgkqhkiG9w0BAQsFAASC
'' SIG '' AgAv/qg4nShD4Dma9yLtT0KPLLdgZm/qSEtG7QtmeWX+
'' SIG '' saCW0fx4dv8SJgyfB4PGSdrybQ1C7stIOhVfCxrURKHj
'' SIG '' lxnP2Mqf5++hMCjjUzjTF2tOhJERI6rD2HsawFJkfE13
'' SIG '' vHVJqIl01GvkgtTIfPn13mNRQt0nTNKTswGUr4urDE11
'' SIG '' DMP1O0kDbUA016zeR7DuQtv9ikg5NT5yf+6CEIls4t+I
'' SIG '' hlf1dzjH9yRIQ+5tQkVxaeegul8yXY1M7TYfPI9Ju+kT
'' SIG '' wk2Devg+TfiO9ytG610kN8/kTiQEPOLZAqQQFNZ/G5Y5
'' SIG '' 0lOi9yAg5/AdzHPbol6Fabvv3yh1xE+ptewwC5OuIdMA
'' SIG '' oGOxn6eYmR6EuVRVuOrUYQ75jwOtOoDbFeRcE8jmjJGk
'' SIG '' In3/aTKq8W5BR+kySKxYhREKxvBCRxLyORJKuRkMnu2R
'' SIG '' GVIVyoWWZtKBck0EKDc+MetQnedXnvUHKt6hL+uIcEdx
'' SIG '' 3p6QZxrLqruAz6Ez4gipFMIxQWJTaNXC3qlX4antTKmv
'' SIG '' yEPFgdXcKaWvFA0x638z5KHEtA/r+iPwGCvzXT/NY0Y+
'' SIG '' AEVkKeNlTmNytzHsWlwegZh7vDImHYfKYc6Orctyeg4+
'' SIG '' OeLRdomVhA0kVzecD8+QKXCPqW7AdiNzA7HjpIbU9Vfw
'' SIG '' vWaPXDqukFm6m73NuEsLQoylkg==
'' SIG '' End signature block

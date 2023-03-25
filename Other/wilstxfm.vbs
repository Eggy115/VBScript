' Windows Installer transform viewer for use with Windows Scripting Host
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the database APIs for viewing transform files
'
Option Explicit

Const iteAddExistingRow      = 1
Const iteDelNonExistingRow   = 2
Const iteAddExistingTable    = 4
Const iteDelNonExistingTable = 8
Const iteUpdNonExistingRow   = 16
Const iteChangeCodePage      = 32
Const iteViewTransform       = 256

Const icdLong       = 0
Const icdShort      = &h400
Const icdObject     = &h800
Const icdString     = &hC00
Const icdNullable   = &h1000
Const icdPrimaryKey = &h2000
Const icdNoNulls    = &h0000
Const icdPersistent = &h0100
Const icdTemporary  = &h0000

Const idoReadOnly = 0

Dim gErrors, installer, base, database, argCount, arg, argValue
gErrors = iteAddExistingRow + iteDelNonExistingRow + iteAddExistingTable + iteDelNonExistingTable + iteUpdNonExistingRow + iteChangeCodePage
Set database = Nothing

' Check arg count, and display help if no all arguments present
argCount = WScript.Arguments.Count
If (argCount < 2) Then
	WScript.Echo "Windows Installer Transform Viewer for Windows Scripting Host (CScript.exe)" &_
		vbNewLine & " 1st non-numeric argument is path to base database which transforms reference" &_
		vbNewLine & " Subsequent non-numeric arguments are paths to the transforms to be viewed" &_
		vbNewLine & " Numeric argument is optional error suppression flags (default is ignore all)" &_
		vbNewLine & " Arguments are executed left-to-right, as encountered" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Cannot run with GUI script host, as listing is performed to standard out
If UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "W" Then
	WScript.Echo "Cannot use WScript.exe - must use CScript.exe with this program"
	Wscript.Quit 2
End If

' Create installer object
On Error Resume Next
Set installer = CreateObject("WindowsInstaller.Installer") : CheckError

' Process arguments, opening database and applying transforms
For arg = 0 To argCount - 1
	argValue = WScript.Arguments(arg)
	If IsNumeric(argValue) Then
		gErrors = argValue
	ElseIf database Is Nothing Then
		Set database = installer.OpenDatabase(argValue, idoReadOnly)
	Else
		database.ApplyTransform argValue, iteViewTransform + gErrors
	End If
	CheckError
Next
ListTransform(database)

Function DecodeColDef(colDef)
	Dim def
	Select Case colDef AND (icdShort OR icdObject)
	Case icdLong
		def = "LONG"
	Case icdShort
		def = "SHORT"
	Case icdObject
		def = "OBJECT"
	Case icdString
		def = "CHAR(" & (colDef AND 255) & ")"
	End Select
	If (colDef AND icdNullable)   =  0 Then def = def & " NOT NULL"
	If (colDef AND icdPrimaryKey) <> 0 Then def = def & " PRIMARY KEY"
	DecodeColDef = def
End Function

Sub ListTransform(database)
	Dim view, record, row, column, change
	On Error Resume Next
	Set view = database.OpenView("SELECT * FROM `_TransformView` ORDER BY `Table`, `Row`") : CheckError
	view.Execute : CheckError
	Do
		Set record = view.Fetch : CheckError
		If record Is Nothing Then Exit Do
		change = Empty
		If record.IsNull(3) Then
			row = "<DDL>"
			If NOT record.IsNull(4) Then change = "[" & record.StringData(5) & "]: " & DecodeColDef(record.StringData(4))
		Else
			row = "[" & Join(Split(record.StringData(3), vbTab, -1), ",") & "]"
			If record.StringData(2) <> "INSERT" AND record.StringData(2) <> "DELETE" Then change = "{" & record.StringData(5) & "}->{" & record.StringData(4) & "}"
		End If
		column = record.StringData(1) & " " & record.StringData(2)
		if Len(column) < 24 Then column = column & Space(24 - Len(column))
		WScript.Echo column, row, change
	Loop
End Sub

Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
	End If
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIIl9gYJKoZIhvcNAQcCoIIl5zCCJeMCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' bSE4J2+6sdFrt3HJs8WbQ0/DQ6WkcvTESWq4fOF4kUmg
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
'' SIG '' Bmrm1MbfI5qWdcUxghnNMIIZyQIBATCBlTB+MQswCQYD
'' SIG '' VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
'' SIG '' A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
'' SIG '' IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
'' SIG '' Q29kZSBTaWduaW5nIFBDQSAyMDEwAhMzAAAEcJ9zF/Wb
'' SIG '' DmRmAAAAAARwMA0GCWCGSAFlAwQCAQUAoIIBBDAZBgkq
'' SIG '' hkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3
'' SIG '' AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQx
'' SIG '' IgQg/M4LDrkBMu8VzYzK+dUbfaQ78LzdU1MwXYrmpbLq
'' SIG '' 3S8wPAYKKwYBBAGCNwoDHDEuDCw5VDF4M2J2UWhPc3F1
'' SIG '' U2w1WlpTUFNyUzlIN2VYK0RUUHVlUVZqSUNUNDBzPTBa
'' SIG '' BgorBgEEAYI3AgEMMUwwSqAkgCIATQBpAGMAcgBvAHMA
'' SIG '' bwBmAHQAIABXAGkAbgBkAG8AdwBzoSKAIGh0dHA6Ly93
'' SIG '' d3cubWljcm9zb2Z0LmNvbS93aW5kb3dzMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAErutcdnghquk/0G8DORd0I+30IGEHEf
'' SIG '' 7k3FOs3sFNkClIwfnZ9moaodEFO19NAqNnvI1AZrgWL5
'' SIG '' IwfV0XtsbT/M210kkZVBQkYryuEjdZFDlR9PhD3Gt8An
'' SIG '' jIB50nmsufl8eXYX1ausFTDrvwUSV9Y+wbQnuCgXOHmA
'' SIG '' hBEvdD3BsPBvQKV2DIRk2xfdKSmUXCWHLI8Jzt8VSODq
'' SIG '' RxrlwfNNeatTTI1A0yDjr4V1t2BdaGwusmxLIy0F2USC
'' SIG '' PmPuedadGU6xl5OUavQee9p3yB4BwypRY6ANSFrZFnq6
'' SIG '' BOQFoab9TQewVBh2KD3iLf8TIyrM/ojjxrwtxuY3YCMx
'' SIG '' Q/mhghcAMIIW/AYKKwYBBAGCNwMDATGCFuwwghboBgkq
'' SIG '' hkiG9w0BBwKgghbZMIIW1QIBAzEPMA0GCWCGSAFlAwQC
'' SIG '' AQUAMIIBUQYLKoZIhvcNAQkQAQSgggFABIIBPDCCATgC
'' SIG '' AQEGCisGAQQBhFkKAwEwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' LIl8A4koJHu1whwstmwq9+SDovkYIaFLCM3YJAdE3CsC
'' SIG '' BmLP9QXS7RgTMjAyMjA3MTYwODU2NTcuNjE2WjAEgAIB
'' SIG '' 9KCB0KSBzTCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgT
'' SIG '' Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
'' SIG '' BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMG
'' SIG '' A1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9u
'' SIG '' czEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046REQ4Qy1F
'' SIG '' MzM3LTJGQUUxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
'' SIG '' LVN0YW1wIFNlcnZpY2WgghFXMIIHDDCCBPSgAwIBAgIT
'' SIG '' MwAAAZwPpk1h0p5LKAABAAABnDANBgkqhkiG9w0BAQsF
'' SIG '' ADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
'' SIG '' Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
'' SIG '' TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1N
'' SIG '' aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0y
'' SIG '' MTEyMDIxOTA1MTlaFw0yMzAyMjgxOTA1MTlaMIHKMQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3Nv
'' SIG '' ZnQgQW1lcmljYSBPcGVyYXRpb25zMSYwJAYDVQQLEx1U
'' SIG '' aGFsZXMgVFNTIEVTTjpERDhDLUUzMzctMkZBRTElMCMG
'' SIG '' A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
'' SIG '' ZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
'' SIG '' ANtSKgwZXUkWP6zrXazTaYq7bco9Q2zvU6MN4ka3GRMX
'' SIG '' 2tJZOK4DxeBiQACL/n7YV/sKTslwpD0f9cPU4rCDX9sf
'' SIG '' cTWo7XPxdHLQ+WkaGbKKWATsqw69bw8hkJ/bjcp2V2A6
'' SIG '' vGsvwcqJCh07BK3JPmUtZikyy5PZ8fyTyiKGN7hOWlaI
'' SIG '' U9oIoucUNoAHQJzLq8h20eNgHUh7eI5k+Kyq4v6810LH
'' SIG '' uA6EHyKJOZN2xTw5JSkLy0FN5Mhg/OaFrFBl3iag2Tqp
'' SIG '' 4InKLt+Jbh/Jd0etnei2aDHFrmlfPmlRSv5wSNX5zAhg
'' SIG '' EyRpjmQcz1zp0QaSAefRkMm923/ngU51IbrVbAeHj569
'' SIG '' SHC9doHgsIxkh0K3lpw582+0ONXcIfIU6nkBT+qADAZ+
'' SIG '' 0dT1uu/gRTBy614QAofjo258TbSX9aOU1SHuAC+3bMoy
'' SIG '' M7jNdHEJROH+msFDBcmJRl4VKsReI5+S69KUGeLIBhhm
'' SIG '' nmQ6drF8Ip0ZiO+vhAsD3e9AnqnY7Hcge850I9oKvwuw
'' SIG '' pVwWnKnwwSGElMz7UvCocmoUMXk7Vn2aNti+bdH28+GQ
'' SIG '' b5EMsqhOmvuZOCRpOWN33G+b3g5unwEP0eTiY+LnWa2A
'' SIG '' uK43z/pplURJVle29K42QPkOcglB6sjLmNpEpb9basJ7
'' SIG '' 2eA0Mlp1LtH3oYZGXsggTfuXAgMBAAGjggE2MIIBMjAd
'' SIG '' BgNVHQ4EFgQUu2kJZ1Ndjl2112SynL6jGMID+rIwHwYD
'' SIG '' VR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYD
'' SIG '' VR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3Nv
'' SIG '' ZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0JTIwVGlt
'' SIG '' ZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsG
'' SIG '' AQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3
'' SIG '' dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNy
'' SIG '' b3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgx
'' SIG '' KS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggr
'' SIG '' BgEFBQcDCDANBgkqhkiG9w0BAQsFAAOCAgEApwAqpiMY
'' SIG '' RzNNYyz3PSbtijbeyCpUXcvIrqA4zPtMIcAk34W9u9mR
'' SIG '' DndWS+tlR3WwTpr1OgaV1wmc6YFzqK6EGWm903UEsFE7
'' SIG '' xBJMPXjfdVOPhcJB3vfvA0PX56oobcF2OvNsOSwTB8bi
'' SIG '' /ns+Cs39Puzs+QSNQZd8iAVBCSvxNCL78dln2RGU1xyB
'' SIG '' 4AKqV9vi4Y/Gfmx2FA+jF0y+YLeob0M40nlSxL0q075t
'' SIG '' 7L6iFRMNr0u8ROhzhDPLl+4ePYfUmyYJoobvydel9anA
'' SIG '' EsHFlhKl+aXb2ic3yNwbsoPycZJL/vo8OVvYYxCy+/5F
'' SIG '' rQmAvoW0ZEaBiYcKkzrNWt/hX9r5KgdwL61x0ZiTZopT
'' SIG '' ko6W/58UTefTbhX7Pni0MApH3Pvyt6N0IFap+/LlwFRD
'' SIG '' 1zn7e6ccPTwESnuo/auCmgPznq80OATA7vufsRZPvqeX
'' SIG '' 8jKtsraSNscvNQymEWlcqdXV9hYkjb4T/Qse9cUYaoXg
'' SIG '' 68wFHFuslWfTdPYPLl1vqzlPMnNJpC8KtdioDgcq+y1B
'' SIG '' aSqSm8EdNfwzT37+/JFtVc3Gs915fDqgPZDgOSzKQIV+
'' SIG '' fw3aPYt2LET3AbmKKW/r13Oy8cg3+D0D362GQBAJVv0N
'' SIG '' RI5NowgaCw6oNgWOFPrN72WSEcca/8QQiTGP2XpLiGpR
'' SIG '' DJZ6sWRpRYNdydkwggdxMIIFWaADAgECAhMzAAAAFcXn
'' SIG '' a54Cm0mZAAAAAAAVMA0GCSqGSIb3DQEBCwUAMIGIMQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3Nv
'' SIG '' ZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkgMjAx
'' SIG '' MDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMyMjVa
'' SIG '' MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5n
'' SIG '' dG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
'' SIG '' aWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1p
'' SIG '' Y3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIICIjAN
'' SIG '' BgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA5OGmTOe0
'' SIG '' ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK
'' SIG '' 4OK9uT4XYDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0T
'' SIG '' z3cywBAY6GB9alKDRLemjkZrBxTzxXb1hlDcwUTIcVxR
'' SIG '' MTegCjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk
'' SIG '' 2pg7uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bs
'' SIG '' hVZN7928jaTjkY+yOSxRnOlwaQ3KNi1wjjHINSi947SH
'' SIG '' JMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedG
'' SIG '' bsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74
'' SIG '' kpEeHT39IM9zfUGaRnXNxF803RKJ1v2lIH1+/NmeRd+2
'' SIG '' ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03dJQc
'' SIG '' NIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5TI4C
'' SIG '' vEJoLhDqhFFG4tG9ahhaYQFzymeiXtcodgLiMxhy16cg
'' SIG '' 8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5UPkLiWHz
'' SIG '' NgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREd
'' SIG '' cu+N+VLEhReTwDwV2xo3xwgVGD94q0W29R6HXtqPnhZy
'' SIG '' acaue7e3PmriLq0CAwEAAaOCAd0wggHZMBIGCSsGAQQB
'' SIG '' gjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5k
'' SIG '' xJq+gpE8RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5i
'' SIG '' XYP05dJlpxtTNRnpcjBcBgNVHSAEVTBTMFEGDCsGAQQB
'' SIG '' gjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3
'' SIG '' Lm1pY3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3Np
'' SIG '' dG9yeS5odG0wEwYDVR0lBAwwCgYIKwYBBQUHAwgwGQYJ
'' SIG '' KwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQD
'' SIG '' AgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU
'' SIG '' 1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0fBE8wTTBL
'' SIG '' oEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3Br
'' SIG '' aS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAt
'' SIG '' MDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEF
'' SIG '' BQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3Br
'' SIG '' aS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5j
'' SIG '' cnQwDQYJKoZIhvcNAQELBQADggIBAJ1VffwqreEsH2cB
'' SIG '' MSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9MTO1OdfCcTY/
'' SIG '' 2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYe
'' SIG '' eNg2LpypglYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB
'' SIG '' 2RjebYIM9W0jVOR4U3UkV7ndn/OOPcbzaN9l9qRWqveV
'' SIG '' tihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSF
'' SIG '' QrP8DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/
'' SIG '' LAl4FOmRsqlb30mjdAy87JGA0j3mSj5mO0+7hvoyGtmW
'' SIG '' 9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1Rt
'' SIG '' nWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPX
'' SIG '' fx5bRAGOWhmRaw2fpCjcZxkoJLo4S5pu+yFUa2pFEUep
'' SIG '' 8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq77EFm
'' SIG '' PWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGConsX
'' SIG '' HRWJjXD+57XQKBqJC4822rpM+Zv/Cuk0+CQ1ZyvgDbjm
'' SIG '' jJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328y+l7vzhw
'' SIG '' RNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUz
'' SIG '' WLOhcGbyoYICzjCCAjcCAQEwgfihgdCkgc0wgcoxCzAJ
'' SIG '' BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAw
'' SIG '' DgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3Nv
'' SIG '' ZnQgQ29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29m
'' SIG '' dCBBbWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRo
'' SIG '' YWxlcyBUU1MgRVNOOkREOEMtRTMzNy0yRkFFMSUwIwYD
'' SIG '' VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
'' SIG '' oiMKAQEwBwYFKw4DAhoDFQDN2Wnq3fCz9ucStub1zQz7
'' SIG '' 129TQKCBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwMA0GCSqGSIb3DQEBBQUAAgUA5nxtwDAiGA8y
'' SIG '' MDIyMDcxNjA2NTA0MFoYDzIwMjIwNzE3MDY1MDQwWjB3
'' SIG '' MD0GCisGAQQBhFkKBAExLzAtMAoCBQDmfG3AAgEAMAoC
'' SIG '' AQACAhR2AgH/MAcCAQACAhGnMAoCBQDmfb9AAgEAMDYG
'' SIG '' CisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAI
'' SIG '' AgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEF
'' SIG '' BQADgYEARKdvMFTyNjPodBnF9sII0hj0YUpZekvsvJ7p
'' SIG '' UbSEfI4aAWhwq7IUyjFzr8L/GJLYHc410q73EW72JEjk
'' SIG '' hGUZhBpFBsVA4XToc9mIhuaF4SHHeNb6PFhM4HCaO6pU
'' SIG '' I2RmLmtiJC4P/5s58IdHfmKmBnvnrN3M0RcwzevYIAOF
'' SIG '' mgIxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzET
'' SIG '' MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
'' SIG '' bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
'' SIG '' aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
'' SIG '' cCBQQ0EgMjAxMAITMwAAAZwPpk1h0p5LKAABAAABnDAN
'' SIG '' BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
'' SIG '' CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCAcmTOU
'' SIG '' L6VHK3YviyU9WF6bN5vdvCTBkcA2E3Zp5vM2jTCB+gYL
'' SIG '' KoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIDcPRYUgjSzK
'' SIG '' OhF39d4QgbRZQgrPO7Lo/qE5GtvSeqa8MIGYMIGApH4w
'' SIG '' fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAGc
'' SIG '' D6ZNYdKeSygAAQAAAZwwIgQggEMpQEe+SDkVnD6hE7l2
'' SIG '' h1gRNE1rq4Oh79tMwbg9E8UwDQYJKoZIhvcNAQELBQAE
'' SIG '' ggIAUfek2fBBYjBJA+/EsAwwgULwj4FymSuhlaPlReWt
'' SIG '' 8U+02zPC6PxxVJ6s+Tqch0Y+sqwSlcrOajjzsz6SgSqd
'' SIG '' f0ThwMP29UVsfnOS2TLeOgNMJdKkIESN0dq3NW1nqdfE
'' SIG '' ZMvrXRwD2R0da6P06Qo19pZpFdo5nvoQUnOqhHbchfVN
'' SIG '' S0geX5U+Ssqy0P4YC604T7LOkqcXQZVGGbQXYJpJb/hZ
'' SIG '' Z5f6vfDDge78/Tb321JtCOjT4/iL6wLuQ3cKRA/N8fbW
'' SIG '' U5Ff2vaZuPovDby5ZvPj7/JbB5GaVTu4JySMkZNrgU/D
'' SIG '' LE4xEQ7ZD8/PSeKXErqVGfC9K7S27TFb2XttUFqNyXdT
'' SIG '' VLgK8ZqFt2zCLoA/a6540UczYYb4rFaAvD7LrVpzPmc4
'' SIG '' 7QOWNwAgaR4L7123Z4W7vV/h7fugxDWq/kvZn+bGmHTd
'' SIG '' iFXQrT9NQMc8TaU7k+700wehhK3c25CKtfN3ysU6Kkk+
'' SIG '' GFDSz7RsieshhNRuYq/3aT0+G86zZk7qUSXvjcq/y+lQ
'' SIG '' IJSwU8OCjOGSDMmAi8uRzF4BRxuP9jPuW/SwyCvZTQxD
'' SIG '' oVnYEk2xVkopcXYahM9C59OQJM64cC7ksB/KGfJzTP5a
'' SIG '' xoZvfE3HX+QCaMW9+TymE9mK6crsn2l4awqzI4myJiVv
'' SIG '' NDMeXveX1wKmiOyHbzcTcAI1xS0=
'' SIG '' End signature block

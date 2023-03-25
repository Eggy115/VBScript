' Windows Installer utility to manage binary streams in an installer package
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the database _Streams table
' Used for entering non-database binary streams such as compressed file cabinets
' Streams that persist database binary values should be managed with table views
' Streams that persist database tables and system data are invisible in _Streams
'
Option Explicit

Const msiOpenDatabaseModeReadOnly     = 0
Const msiOpenDatabaseModeTransact     = 1
Const msiOpenDatabaseModeCreate       = 3

Const msiViewModifyInsert         = 1
Const msiViewModifyUpdate         = 2
Const msiViewModifyAssign         = 3
Const msiViewModifyReplace        = 4
Const msiViewModifyDelete         = 6

Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
Const TristateTrue = -1

' Check arg count, and display help if argument not present or contains ?
Dim argCount:argCount = Wscript.Arguments.Count
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If (argCount = 0) Then
	Wscript.Echo "Windows Installer database stream import utility" &_
		vbNewLine & " 1st argument is the path to MSI database (installer package)" &_
		vbNewLine & " 2nd argument is the path to a file containing the stream data" &_
		vbNewLine & " If the 2nd argument is missing, streams will be listed" &_
		vbNewLine & " 3rd argument is optional, the name used for the stream" &_
		vbNewLine & " If the 3rd arugment is missing, the file name is used" &_
		vbNewLine & " To remove a stream, use /D or -D as the 2nd argument" &_
		vbNewLine & " followed by the name of the stream in the 3rd argument" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

' Evaluate command-line arguments and set open and update modes
Dim databasePath:databasePath = Wscript.Arguments(0)
Dim openMode    : If argCount = 1 Then openMode = msiOpenDatabaseModeReadOnly Else openMode = msiOpenDatabaseModeTransact
Dim updateMode  : If argCount > 1 Then updateMode = msiViewModifyAssign  'Either insert or replace existing row
Dim importPath  : If argCount > 1 Then importPath = Wscript.Arguments(1)
Dim streamName  : If argCount > 2 Then streamName = Wscript.Arguments(2)
If streamName = Empty And importPath <> Empty Then streamName = Right(importPath, Len(importPath) - InStrRev(importPath, "\",-1,vbTextCompare))
If UCase(importPath) = "/D" Or UCase(importPath) = "-D" Then updateMode = msiViewModifyDelete : importPath = Empty 'Stream will be deleted if no input data

' Open database and create a view on the _Streams table
Dim sqlQuery : Select Case updateMode
	Case msiOpenDatabaseModeReadOnly: sqlQuery = "SELECT `Name` FROM _Streams"
	Case msiViewModifyAssign:         sqlQuery = "SELECT `Name`,`Data` FROM _Streams"
	Case msiViewModifyDelete:         sqlQuery = "SELECT `Name` FROM _Streams WHERE `Name` = ?"
End Select
Dim database : Set database = installer.OpenDatabase(databasePath, openMode) : CheckError
Dim view     : Set view = database.OpenView(sqlQuery)
Dim record

If openMode = msiOpenDatabaseModeReadOnly Then 'If listing streams, simply fetch all records
	Dim message, name
	view.Execute : CheckError
	Do
		Set record = view.Fetch
		If record Is Nothing Then Exit Do
		name = record.StringData(1)
		If message = Empty Then message = name Else message = message & vbNewLine & name
	Loop
	Wscript.Echo message
Else 'If adding a stream, insert a row, else if removing a stream, delete the row
	Set record = installer.CreateRecord(2)
	record.StringData(1) = streamName
	view.Execute record : CheckError
	If importPath <> Empty Then  'Insert stream - copy data into stream
		record.SetStream 2, importPath : CheckError
	Else  'Delete stream, fetch first to provide better error message if missing
		Set record = view.Fetch
		If record Is Nothing Then Wscript.Echo "Stream not present:", streamName : Wscript.Quit 2
	End If
	view.Modify updateMode, record : CheckError
	database.Commit : CheckError
	Set view = Nothing
	Set database = Nothing
	CheckError
End If

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
'' SIG '' cI6yPc6nU50pwvCvjhpQqcdV9XZAulr8Q0FGvLPiuv6g
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
'' SIG '' IgQgQokwC+HG3/TZ4gbQnFv4cBB6+FknILF1a4Hcmtoi
'' SIG '' sjAwPAYKKwYBBAGCNwoDHDEuDCxpc0xPVUN5a0NDbGZ3
'' SIG '' dUdPNWkwbXhZblo4WTQ3UE56RWFPTjBPbkM2WjRJPTBa
'' SIG '' BgorBgEEAYI3AgEMMUwwSqAkgCIATQBpAGMAcgBvAHMA
'' SIG '' bwBmAHQAIABXAGkAbgBkAG8AdwBzoSKAIGh0dHA6Ly93
'' SIG '' d3cubWljcm9zb2Z0LmNvbS93aW5kb3dzMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAFTMvVYUYeUe4AAgeTqeGD4ZmSUYiM25
'' SIG '' VFgnBrKTr6RJ0aQtfo9jfXd93fornR6MQ+2VswDWAJby
'' SIG '' vPRmliH9h5RxEU0Ve4erG2BQBgjiBfW9kjXqtAuI5uTt
'' SIG '' /QdH7ZEKJcjKCu+aFqebH0iz7LSY3SCpIsTzXOOj9OR1
'' SIG '' SKVjUAyAO2+nuwz7aKMsuOQoQVnn3ruxfd+ikHAlkYnO
'' SIG '' nORJ2367EMeM+9ZLQIEtpIzs44yKmKk8AC4sf5qzly/a
'' SIG '' +orVBwZ+D12/vPwLG5z1aKgFrSFrBTNV8YFlH0Z5aBKS
'' SIG '' P4rTzIbXhrsd7zF+HEgl/+P6c0GvZJT33Lcq05Mhy8UT
'' SIG '' 4kqhghcAMIIW/AYKKwYBBAGCNwMDATGCFuwwghboBgkq
'' SIG '' hkiG9w0BBwKgghbZMIIW1QIBAzEPMA0GCWCGSAFlAwQC
'' SIG '' AQUAMIIBUQYLKoZIhvcNAQkQAQSgggFABIIBPDCCATgC
'' SIG '' AQEGCisGAQQBhFkKAwEwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' xzlx2FllZ2jXCIPc8rnhkykVVyhNjiRsEup+zP3BsbQC
'' SIG '' BmK0ysm2YBgTMjAyMjA3MTYwODU3MTAuNTY1WjAEgAIB
'' SIG '' 9KCB0KSBzTCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgT
'' SIG '' Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
'' SIG '' BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjElMCMG
'' SIG '' A1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9u
'' SIG '' czEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046MjI2NC1F
'' SIG '' MzNFLTc4MEMxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1l
'' SIG '' LVN0YW1wIFNlcnZpY2WgghFXMIIHDDCCBPSgAwIBAgIT
'' SIG '' MwAAAZh2s4zF0AWhAQABAAABmDANBgkqhkiG9w0BAQsF
'' SIG '' ADB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
'' SIG '' Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMV
'' SIG '' TWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1N
'' SIG '' aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0y
'' SIG '' MTEyMDIxOTA1MTVaFw0yMzAyMjgxOTA1MTVaMIHKMQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3Nv
'' SIG '' ZnQgQW1lcmljYSBPcGVyYXRpb25zMSYwJAYDVQQLEx1U
'' SIG '' aGFsZXMgVFNTIEVTTjoyMjY0LUUzM0UtNzgwQzElMCMG
'' SIG '' A1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
'' SIG '' ZTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIB
'' SIG '' AMbUlaxWSynzEbiwsyd/F+K3dKEj7sbUx9NP7le9DO4A
'' SIG '' 57yvkxEAUhNOaMXHOgsV+ZrEu89WWYOCQOLSuqw6z0CX
'' SIG '' 2NXBhIVUX/BYLb4Hvo7KyLJGPD40+PkDhyYyE+oh02RE
'' SIG '' sIT7C24j/AJqrf8t/iSgMa50hwRhGAyqpOg45QhXh7sR
'' SIG '' 1hveT2tg83tKyXCwsVKn4W+b9BzLkqp+SYxfhLegnHsd
'' SIG '' 2JCEpsrULpl+Jv7vrVuat08tPp512WfLCWzuEKsgi4W2
'' SIG '' BRtSPookhmfUxthjyGsAzn228ul4aYVbcaN4ECa8HECf
'' SIG '' uj0unafKRPXD0jSz113CkWeMtPY8rvgYNKzEVRkbVS0v
'' SIG '' KmL+RlyD1Z6c8BmlS08V87ky2J/wlryNdcsg/or5vkuJ
'' SIG '' BXygjEVIF+AU3v9Mva1JJ9BVy+pfWZxI6vH+2yCrcvpg
'' SIG '' DEjo+XiHXNCtwCZOjKkSg9g1z9GVIGTqWOY3I0OxfeC0
'' SIG '' rynpzscJZSEX5iMyB9qdCYyNRixuN0SwLIvpACiNnR/q
'' SIG '' S143hxXqhsXBxQS+JjKBZt51pPzo4Z70sQ7E+6HOAW/Z
'' SIG '' mhtWvQnyGXUVV1xkVt8U3+B2Mdn+dwMOos1aBygygSHD
'' SIG '' DOjsUA5uoprF8HnMIGphKPjmaI07mDeE/wCALR5IIeXe
'' SIG '' srsk8yvUH7wlMe3BGRIrP/5zAgMBAAGjggE2MIIBMjAd
'' SIG '' BgNVHQ4EFgQUbpGEco2myDeaCiezstHlgdPN4TcwHwYD
'' SIG '' VR0jBBgwFoAUn6cVXQBeYl2D9OXSZacbUzUZ6XIwXwYD
'' SIG '' VR0fBFgwVjBUoFKgUIZOaHR0cDovL3d3dy5taWNyb3Nv
'' SIG '' ZnQuY29tL3BraW9wcy9jcmwvTWljcm9zb2Z0JTIwVGlt
'' SIG '' ZS1TdGFtcCUyMFBDQSUyMDIwMTAoMSkuY3JsMGwGCCsG
'' SIG '' AQUFBwEBBGAwXjBcBggrBgEFBQcwAoZQaHR0cDovL3d3
'' SIG '' dy5taWNyb3NvZnQuY29tL3BraW9wcy9jZXJ0cy9NaWNy
'' SIG '' b3NvZnQlMjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgx
'' SIG '' KS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggr
'' SIG '' BgEFBQcDCDANBgkqhkiG9w0BAQsFAAOCAgEAJPoHoXfe
'' SIG '' L/z3NdOCpDwvoJgwfH0GJoc5X7CTnck6uILN5ouNiBHK
'' SIG '' ywmGaecn8J0drmqNxLC9Gm1alkk9UrmzGE4iNEE+Cz/f
'' SIG '' 4RHS9LzsgD5oZt/s0XstlmXFY86X/IUGD2pne2k4Y6iF
'' SIG '' AidCfnOlXbeFailo3hzj2MYkcs8B/L27v5lIZC7DXgKx
'' SIG '' b9dEsQsdPXwjrRbS4o4Frk+bZWKiEyi9xuk1QIQRGog7
'' SIG '' 1Y/DMjAxFHDfj8uCO6yUcmin7/VV78J/I2rB5SbB6lAc
'' SIG '' mt37BMtSWCbgQ1tcXqLnaMV9ikRLAt0Cfnqj+mP6Cux3
'' SIG '' YusAQ9BHKHj2ta8j+pl86G1PYVabMXDogm9nsLNPU74V
'' SIG '' zSAgME2pqyzlBuaQ6QpjL1TucUDqqfdln4ytkywlOPuD
'' SIG '' EB/TIyRWrBhZlGThutj2rwkM+Zx81KNGtV+ljLMRUSp6
'' SIG '' YZqebG8MNPNLbCRIFrfNw3A6BiFYFOYl0uDKJYkZ6rKP
'' SIG '' WblvA2Cc7Do3NcKJUzN9vO12So51NHzwu0AkY1GN69aN
'' SIG '' B3leK0a56BKnaYwmCUXNHCSdxBq7UEmwKP/VoNjigyI7
'' SIG '' xyieSZpYGth7XVAJLz3r+xnBJ2cRQlqTSqmcFEUH5MdE
'' SIG '' jEiK8Io1vEbZBFnx2H3lw5eCjRi8E3lrWn6Ine83DOd5
'' SIG '' TYAgLvPeushs3Z8wggdxMIIFWaADAgECAhMzAAAAFcXn
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
'' SIG '' YWxlcyBUU1MgRVNOOjIyNjQtRTMzRS03ODBDMSUwIwYD
'' SIG '' VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNl
'' SIG '' oiMKAQEwBwYFKw4DAhoDFQDzLB7+IXkzx8hTZpPrJDe+
'' SIG '' c+lXk6CBgzCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwMA0GCSqGSIb3DQEBBQUAAgUA5nzxwzAiGA8y
'' SIG '' MDIyMDcxNjE2MTM1NVoYDzIwMjIwNzE3MTYxMzU1WjB3
'' SIG '' MD0GCisGAQQBhFkKBAExLzAtMAoCBQDmfPHDAgEAMAoC
'' SIG '' AQACAg2jAgH/MAcCAQACAhGhMAoCBQDmfkNDAgEAMDYG
'' SIG '' CisGAQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAI
'' SIG '' AgEAAgMHoSChCjAIAgEAAgMBhqAwDQYJKoZIhvcNAQEF
'' SIG '' BQADgYEAm9iUR+Kx1geqSIQRIU7SQtyJt/Ou3n4yNfgw
'' SIG '' aF6EfViguw+3Rh+m0WGxAGpakbBthK/Zbd6zaDDzI5M5
'' SIG '' PtTFVLPJWo23T2s8BoVEcy6PAeR2y7W2JbAwi+Rli2Et
'' SIG '' mFEjgEh25iq+4PlzJmvygaGBsGtaX2qq1a9lKDPIwhuR
'' SIG '' sooxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzET
'' SIG '' MBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
'' SIG '' bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
'' SIG '' aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFt
'' SIG '' cCBQQ0EgMjAxMAITMwAAAZh2s4zF0AWhAQABAAABmDAN
'' SIG '' BglghkgBZQMEAgEFAKCCAUowGgYJKoZIhvcNAQkDMQ0G
'' SIG '' CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCBCZG5h
'' SIG '' CwgCxAeDro59wyjgqDn9aEZq13ZLkw2JnEniKDCB+gYL
'' SIG '' KoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIL+mzgY5Of/3
'' SIG '' A7U2Ecz1B97SWgHeyWTDUUXev5uHbVbEMIGYMIGApH4w
'' SIG '' fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAGY
'' SIG '' drOMxdAFoQEAAQAAAZgwIgQgexxmgbCyvEYx2zF0cGWs
'' SIG '' xqTPf7yAaPNaTa+ExvkMZcAwDQYJKoZIhvcNAQELBQAE
'' SIG '' ggIAiBdzdswKeENhJsVoAwardB6p7wRYJVq7hruBw37H
'' SIG '' TlOLobqXNUnVryjEwFvkrcW86KncdLi+0kanI2w6SuOf
'' SIG '' Kl6GuB6wOpo/55iI2Vy4nsuyvH6BnYwU56xL4fa0Lysc
'' SIG '' 9El+vYI9EqRm6RKtq1x/DZUo/6uZP3gzLE70/iDqUcvA
'' SIG '' 3C4t3Je013YN60rRjQlgFISQgmeMr42TLBpXeTYfO1yC
'' SIG '' gIU13uYYqRoto9yN9sRTP1LJHDNoUpPUAwUNhSMStxjD
'' SIG '' TMU5GtEEx5OSWdmyIKS/3YVwg+9QcQ1rp/zNJrAETAcR
'' SIG '' QKRdTRLt55anPlvBLI00BrtzA3JV793/9XkQlpLUD7C9
'' SIG '' fg4RbA8LpVE34q6Le4MSItTpphDnc2p6+K8CGhZj7GkS
'' SIG '' cYFIPiQw1EWEC9UShosLMDKOjo+QFkoNKwNkQDg4vwOZ
'' SIG '' CVD1f/m606TySCoDrpE62ONUmNFM60E8u6h8qyxUSZTn
'' SIG '' kUEdh5norqAo3K/S1RqEjAw6PIQOVz0Fg2uZKws78gin
'' SIG '' xwxkdW2Puh6B75K3JyITjCTClXW+H/yf/SLfyV6XAiet
'' SIG '' E+mfgGLOP4iPsRuqrIjH47r110P2+UZ83Na2xawwIZCl
'' SIG '' swxA4F2DJRKzMUE6TAuS1Vsl+uj/6SrCsacpR47uNN4h
'' SIG '' Kl6GuO2HAvOL7OOYHigc8gfAXIk=
'' SIG '' End signature block

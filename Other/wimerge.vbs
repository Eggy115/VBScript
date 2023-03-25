' Windows Installer database utility to merge data from another database              
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the Database.Merge method and MsiDatabaseMerge API
'
Option Explicit

Const msiOpenDatabaseModeReadOnly     = 0
Const msiOpenDatabaseModeTransact     = 1
Const msiOpenDatabaseModeCreate       = 3
Const ForAppending = 8
Const ForReading = 1
Const ForWriting = 2
Const TristateTrue = -1

Dim argCount:argCount = Wscript.Arguments.Count
Dim iArg:iArg = 0
If (argCount < 2) Then
	Wscript.Echo "Windows Installer database merge utility" &_
		vbNewLine & " 1st argument is the path to MSI database (installer package)" &_
		vbNewLine & " 2nd argument is the path to database containing data to merge" &_
		vbNewLine & " 3rd argument is the optional table to contain the merge errors" &_
		vbNewLine & " If 3rd argument is not present, the table _MergeErrors is used" &_
		vbNewLine & "  and that table will be dropped after displaying its contents." &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

' Open databases and merge data
Dim database1 : Set database1 = installer.OpenDatabase(WScript.Arguments(0), msiOpenDatabaseModeTransact) : CheckError
Dim database2 : Set database2 = installer.OpenDatabase(WScript.Arguments(1), msiOpenDatabaseModeReadOnly) : CheckError
Dim errorTable : errorTable = "_MergeErrors"
If argCount >= 3 Then errorTable = WScript.Arguments(2)
Dim hasConflicts:hasConflicts = database1.Merge(database2, errorTable) 'Old code returns void value, new returns boolean
If hasConflicts <> True Then hasConflicts = CheckError 'Temp for old Merge function that returns void
If hasConflicts <> 0 Then
	Dim message, line, view, record
	Set view = database1.OpenView("Select * FROM `" & errorTable & "`") : CheckError
	view.Execute
	Do
		Set record = view.Fetch
		If record Is Nothing Then Exit Do
		line = record.StringData(1) & " table has " & record.IntegerData(2) & " conflicts"
		If message = Empty Then message = line Else message = message & vbNewLine & line
	Loop
	Set view = Nothing
	Wscript.Echo message
End If
If argCount < 3 And hasConflicts Then database1.OpenView("DROP TABLE `" & errorTable & "`").Execute : CheckError
database1.Commit : CheckError
Quit 0

Function CheckError
	Dim message, errRec
	CheckError = 0
	If Err = 0 Then Exit Function
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText : CheckError = errRec.IntegerData(1)
	End If
	If CheckError = 2268 Then Err.Clear : Exit Function
	Wscript.Echo message
	Wscript.Quit 2
End Function

'' SIG '' Begin signature block
'' SIG '' MIIl8gYJKoZIhvcNAQcCoIIl4zCCJd8CAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' QXX+BeRpnj5/3w9MZiLTEbzssoFPyxBqr0/6QcQWjb+g
'' SIG '' ggt9MIIFBTCCA+2gAwIBAgITMwAABG8vaU5Sum9NZAAA
'' SIG '' AAAEbzANBgkqhkiG9w0BAQsFADB+MQswCQYDVQQGEwJV
'' SIG '' UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
'' SIG '' UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
'' SIG '' cmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQgQ29kZSBT
'' SIG '' aWduaW5nIFBDQSAyMDEwMB4XDTIyMDEyNzE5MzIyMFoX
'' SIG '' DTIzMDEyNjE5MzIyMFowfzELMAkGA1UEBhMCVVMxEzAR
'' SIG '' BgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
'' SIG '' bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
'' SIG '' bjEpMCcGA1UEAxMgTWljcm9zb2Z0IFdpbmRvd3MgS2l0
'' SIG '' cyBQdWJsaXNoZXIwggEiMA0GCSqGSIb3DQEBAQUAA4IB
'' SIG '' DwAwggEKAoIBAQCkYtfwfzH8DK2hkGgu/I8/9dcRlY/1
'' SIG '' EA0AYE4OrXcNmTcrpcHtMon4d5O+FPHYH8pq/lIXeemc
'' SIG '' vB1oN28H+VqyXed2R/PLA/UWQ3tEpNPx1t6yMI3wEa9c
'' SIG '' t4UICs3fttUwhQmIchx2APVG+OqmFbdSv/M75KXdYVIO
'' SIG '' 70XUhdibsBcllOS7ySIwc7w4nak8SxxuEF9GF3AgkLLs
'' SIG '' 2md7J3ZEX17dc8TTeGDEvwZ1C8cwHT7WCPLjecSNGS3/
'' SIG '' u2lRouLB9cebw+cnpS+KW6OdbtuWFhJN06LO5DAgg6aZ
'' SIG '' epbNEf927sNUFcooQZtsW4NFM+gjpM0s7G7kBDzronTk
'' SIG '' vfxbAgMBAAGjggF5MIIBdTAfBgNVHSUEGDAWBgorBgEE
'' SIG '' AYI3CgMUBggrBgEFBQcDAzAdBgNVHQ4EFgQUK+ldBYs2
'' SIG '' valv07G9G+lUixKkscwwUAYDVR0RBEkwR6RFMEMxKTAn
'' SIG '' BgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1ZXJ0
'' SIG '' byBSaWNvMRYwFAYDVQQFEw0yMjk5MDMrNDY5MDYxMB8G
'' SIG '' A1UdIwQYMBaAFOb8X3u7IgBY5HJOtfQhdCMy5u+sMFYG
'' SIG '' A1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9z
'' SIG '' b2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNp
'' SIG '' Z1BDQV8yMDEwLTA3LTA2LmNybDBaBggrBgEFBQcBAQRO
'' SIG '' MEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWljcm9z
'' SIG '' b2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzIw
'' SIG '' MTAtMDctMDYuY3J0MAwGA1UdEwEB/wQCMAAwDQYJKoZI
'' SIG '' hvcNAQELBQADggEBAFIbJYffEpkOU3PyGTw3ToSHyUL+
'' SIG '' yIarPZ5b4jaJHM2/fp6W58b+Y0KksQI+cyeYUJ/YjBIt
'' SIG '' 0KEztgmnoN56IoG1OeekBT/Zh53T8fE+TIZHjO6D7scY
'' SIG '' ETENGr3grEcDFNy8zVZPo4DrXWPJt5IKq+Tn9Q2Asf53
'' SIG '' Mq0sunZf3q6VV6tsmzgTCgixPYeh0pSKGqJit2f9jBho
'' SIG '' QbDXcQ1TUjQ37ea7rh4CSEuKMfUdPaHt/C2lCY/YZcxD
'' SIG '' z41o0OjLUgVArAkL5jF6KZtWauWMHgjRGEhS9MMk/FgO
'' SIG '' JbkxHJA6RSto5m/ujxakGBAJk9HM/81KZ/RnQBlTe8h1
'' SIG '' ReBerbQwggZwMIIEWKADAgECAgphDFJMAAAAAAADMA0G
'' SIG '' CSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEG
'' SIG '' A1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
'' SIG '' ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
'' SIG '' MTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZp
'' SIG '' Y2F0ZSBBdXRob3JpdHkgMjAxMDAeFw0xMDA3MDYyMDQw
'' SIG '' MTdaFw0yNTA3MDYyMDUwMTdaMH4xCzAJBgNVBAYTAlVT
'' SIG '' MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
'' SIG '' ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9y
'' SIG '' YXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
'' SIG '' Z25pbmcgUENBIDIwMTAwggEiMA0GCSqGSIb3DQEBAQUA
'' SIG '' A4IBDwAwggEKAoIBAQDpDmRQeWe1xOP9CQBMnpSs91Zo
'' SIG '' 6kTYz8VYT6mldnxtRbrTOZK0pB75+WWC5BfSj/1EnAjo
'' SIG '' ZZPOLFWEv30I4y4rqEErGLeiS25JTGsVB97R0sKJHnGU
'' SIG '' zbV/S7SvCNjMiNZrF5Q6k84mP+zm/jSYV9UdXUn2siou
'' SIG '' 1YW7WT/4kLQrg3TKK7M7RuPwRknBF2ZUyRy9HcRVYldy
'' SIG '' +Ge5JSA03l2mpZVeqyiAzdWynuUDtWPTshTIwciKJgpZ
'' SIG '' fwfs/w7tgBI1TBKmvlJb9aba4IsLSHfWhUfVELnG6Kru
'' SIG '' i2otBVxgxrQqW5wjHF9F4xoUHm83yxkzgGqJTaNqZmN4
'' SIG '' k9Uwz5UfAgMBAAGjggHjMIIB3zAQBgkrBgEEAYI3FQEE
'' SIG '' AwIBADAdBgNVHQ4EFgQU5vxfe7siAFjkck619CF0IzLm
'' SIG '' 76wwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYD
'' SIG '' VR0PBAQDAgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0j
'' SIG '' BBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYDVR0f
'' SIG '' BE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQu
'' SIG '' Y29tL3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0
'' SIG '' XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEBBE4wTDBK
'' SIG '' BggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQu
'' SIG '' Y29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0w
'' SIG '' Ni0yMy5jcnQwgZ0GA1UdIASBlTCBkjCBjwYJKwYBBAGC
'' SIG '' Ny4DMIGBMD0GCCsGAQUFBwIBFjFodHRwOi8vd3d3Lm1p
'' SIG '' Y3Jvc29mdC5jb20vUEtJL2RvY3MvQ1BTL2RlZmF1bHQu
'' SIG '' aHRtMEAGCCsGAQUFBwICMDQeMiAdAEwAZQBnAGEAbABf
'' SIG '' AFAAbwBsAGkAYwB5AF8AUwB0AGEAdABlAG0AZQBuAHQA
'' SIG '' LiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAadO9XTyl7xBaF
'' SIG '' eLhQ0yL8CZ2sgpf4NP8qLJeVEuXkv8+/k8jjNKnbgbjc
'' SIG '' HgC+0jVvr+V/eZV35QLU8evYzU4eG2GiwlojGvCMqGJR
'' SIG '' RWcI4z88HpP4MIUXyDlAptcOsyEp5aWhaYwik8x0mOeh
'' SIG '' R0PyU6zADzBpf/7SJSBtb2HT3wfV2XIALGmGdj1R26Y5
'' SIG '' SMk3YW0H3VMZy6fWYcK/4oOrD+Brm5XWfShRsIlKUaSa
'' SIG '' bMi3H0oaDmmp19zBftFJcKq2rbtyR2MX+qbWoqaG7KgQ
'' SIG '' RJtjtrJpiQbHRoZ6GD/oxR0h1Xv5AiMtxUHLvx1MyBbv
'' SIG '' sZx//CJLSYpuFeOmf3Zb0VN5kYWd1dLbPXM18zyuVLJS
'' SIG '' R2rAqhOV0o4R2plnXjKM+zeF0dx1hZyHxlpXhcK/3Q2P
'' SIG '' jJst67TuzyfTtV5p+qQWBAGnJGdzz01Ptt4FVpd69+lS
'' SIG '' TfR3BU+FxtgL8Y7tQgnRDXbjI1Z4IiY2vsqxjG6qHeSF
'' SIG '' 2kczYo+kyZEzX3EeQK+YZcki6EIhJYocLWDZN4lBiSoW
'' SIG '' D9dhPJRoYFLv1keZoIBA7hWBdz6c4FMYGlAdOJWbHmYz
'' SIG '' Eyc5F3iHNs5Ow1+y9T1HU7bg5dsLYT0q15IszjdaPkBC
'' SIG '' MaQfEAjCVpy/JF1RAp1qedIX09rBlI4HeyVxRKsGaubU
'' SIG '' xt8jmpZ1xTGCGc0wghnJAgEBMIGVMH4xCzAJBgNVBAYT
'' SIG '' AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
'' SIG '' EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
'' SIG '' cG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2Rl
'' SIG '' IFNpZ25pbmcgUENBIDIwMTACEzMAAARvL2lOUrpvTWQA
'' SIG '' AAAABG8wDQYJYIZIAWUDBAIBBQCgggEEMBkGCSqGSIb3
'' SIG '' DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
'' SIG '' DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCB9
'' SIG '' ZNrAIUSORGzQaD3sArrVgPFTTGrK3FiVXZNuVgrDzzA8
'' SIG '' BgorBgEEAYI3CgMcMS4MLENvczVXOVpZTFFOV2lISXVv
'' SIG '' RE9uNXFiU2ptUVpqY3ZtNWxjd1FSU3E4Qkk9MFoGCisG
'' SIG '' AQQBgjcCAQwxTDBKoCSAIgBNAGkAYwByAG8AcwBvAGYA
'' SIG '' dAAgAFcAaQBuAGQAbwB3AHOhIoAgaHR0cDovL3d3dy5t
'' SIG '' aWNyb3NvZnQuY29tL3dpbmRvd3MwDQYJKoZIhvcNAQEB
'' SIG '' BQAEggEAI9ZzSU2DtBGCitsKKpWhxthQiQjUzuPu0YVN
'' SIG '' gXvThuNDoM/1s9nPN4+nlIY35wxC/46Wi8v9A7p8qTwA
'' SIG '' cCKIxiyVTQkc5Ixt6Z6N52K2K36Lg7p5LGaZ+M7FAKJk
'' SIG '' S3peYN9yiOEcK0A29vuCxtHA+3EMHywvURD8Pv34I7Fd
'' SIG '' qpoAHb+DtgDRYyjD3uQ66Mzcn5060Jx5oHVpoHR3HboQ
'' SIG '' tNgJReGEoow5a2HebQ5mctGhs0YCoiptNDt1Tuz9O/rt
'' SIG '' Q9aNzrvbns8eRx4MlNX93sDpKkKJSFOU8T/MlyHjMxb0
'' SIG '' PRKRLdVOEwTjCvRYXv6+AdrfNZHGC2inP6jMTy6p26GC
'' SIG '' FwAwghb8BgorBgEEAYI3AwMBMYIW7DCCFugGCSqGSIb3
'' SIG '' DQEHAqCCFtkwghbVAgEDMQ8wDQYJYIZIAWUDBAIBBQAw
'' SIG '' ggFRBgsqhkiG9w0BCRABBKCCAUAEggE8MIIBOAIBAQYK
'' SIG '' KwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCDe+aqe
'' SIG '' x/Bwm0k38Lr6T8McOdjXfy4yxQ9/7kYo3YrJrAIGYrSf
'' SIG '' u9/pGBMyMDIyMDcxNjA4NTY1Ny44NTNaMASAAgH0oIHQ
'' SIG '' pIHNMIHKMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
'' SIG '' aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
'' SIG '' ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
'' SIG '' ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMSYw
'' SIG '' JAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpFNUE2LUUyN0Mt
'' SIG '' NTkyRTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3Rh
'' SIG '' bXAgU2VydmljZaCCEVcwggcMMIIE9KADAgECAhMzAAAB
'' SIG '' lbf8DdbjNzElAAEAAAGVMA0GCSqGSIb3DQEBCwUAMHwx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
'' SIG '' c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTIxMTIw
'' SIG '' MjE5MDUxMloXDTIzMDIyODE5MDUxMlowgcoxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
'' SIG '' bWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxl
'' SIG '' cyBUU1MgRVNOOkU1QTYtRTI3Qy01OTJFMSUwIwYDVQQD
'' SIG '' ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIIC
'' SIG '' IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEAn21B
'' SIG '' DGe2Szs/WqEQniS+IYU/UPCWQdsWlZTDQrd28IXEyORi
'' SIG '' z67dnvdwwLJpajs8NXBYjz4OkubCwl8+y221EKS4WvEu
'' SIG '' L9qnHDLU6JBGg0EvkCRK5wLJelUpkbwMtJ5Y/gvz2mbi
'' SIG '' 29zs2NAEcO1HgmS6cljzx/pOTHWI+jVA0zaF6m80Bwrj
'' SIG '' 7Pn4CKK6Octwx6DtO+4OiK9kxyMdcn1RRLecw3BTzmDI
'' SIG '' OMgYuAOl3N4ZvbWesPOPZwb1SsJuWAC3x98v395+C5ze
'' SIG '' tW9cMwMd2QmY39d1Cm6RO6eg2Cax0Qf/qcBYxvfU8Bx+
'' SIG '' rl8w3mU+v6+qh+wAAcJ/H6WHNU5pXhWPGEblc846fVZD
'' SIG '' x1fFc78yy+0CtpLXnlyy/2OJb4y+oc8jphPtS1Q95RG2
'' SIG '' IaNcwrfhe21PhaY8gX0wuIv8B7KbW9tfGJW5ELdYtQep
'' SIG '' ZZicFRcAi1+4MUOPECBlGnDMvJKdfs3M2PksZaWhIDZk
'' SIG '' JH3Na2j4fcubDGul+PPsdCuwfDqg6F3E4hAiIyXrccLb
'' SIG '' gZULHidOR0X4rH4BZtPZBu73RxKNzW1LjDARYpHOG6Df
'' SIG '' VH5tIlIavybaldCsK7/Qr92sg4HTcBFoi9muuSJxFkqU
'' SIG '' U2H7AkNN3qhIeQN68Ffyn1BXIrfg6z/vVXA6Y1kbAqJG
'' SIG '' b+LYJ+agFzTLR2vDYqkCAwEAAaOCATYwggEyMB0GA1Ud
'' SIG '' DgQWBBSrl9NiAhRXV4K3AgZgyXx+b/ypFzAfBgNVHSME
'' SIG '' GDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8E
'' SIG '' WDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
'' SIG '' b20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0
'' SIG '' YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUH
'' SIG '' AQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1p
'' SIG '' Y3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29m
'' SIG '' dCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
'' SIG '' dDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsGAQUF
'' SIG '' BwMIMA0GCSqGSIb3DQEBCwUAA4ICAQDgszbeHyfozr0L
'' SIG '' qtCLZ9+yGa2DQRrMAIviABTN2Biv8BkJRJ3II5jQbmnP
'' SIG '' eVtnwC+sbRVXzH5HqkizC6qInVbFPQZuAxAY2ljTk/bl
'' SIG '' /7XGIiUnxUDNKw265fFeJzPPEWReehv6iVvYOXSKjkqI
'' SIG '' psylLf0O1h+lQcltLGq+cBr4KLyt6hWncCkoc0WHBKk5
'' SIG '' Bx9s4qeXu943szx8dvrWmKiRucSc3QxK2dZzIsUY2h7N
'' SIG '' yqXLJmWLsbCEXwWDibwBRspkxkb+T7sLDabPRHIdQGrK
'' SIG '' vOB/2P/MTdxkI+D9zIg5/Is1AQwrlyHx2JN/W6p2gJhW
'' SIG '' 1Igm8vllqbs3ZOKAys/7FsK57KEO9rhBlRDe/pMsPfh0
'' SIG '' qOYvJfGYNWJo/bVIA6VVBowHbqC8h0O16pJypkvZCUgS
'' SIG '' pOKJBA4NCHei3ii0MB9XuGlXk8lGMHAV98IO6SyUFr0e
'' SIG '' 52tkhq7Zf9t2BkE7nZljq8ocfZZ1OygRlf2jb89LU6XL
'' SIG '' LnLCvnGRSgxJFgf6FBVa7crp+jQ+aWGTY9AoEbqeYK1Q
'' SIG '' AqvwIG/hDhiwg/sxLRjaKeLXyr7GG+uNuezSfV6zB4KQ
'' SIG '' om++lk9+ET5ggQcsS1JB8R6ucwsmDbtCBVwLdQFYnMNe
'' SIG '' DPnMy2CFTOzTslaRXXAdQfTIiYpO6XkootF00XZef1fy
'' SIG '' rHE2ggRc9zCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
'' SIG '' SZkAAAAAABUwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBS
'' SIG '' b290IENlcnRpZmljYXRlIEF1dGhvcml0eSAyMDEwMB4X
'' SIG '' DTIxMDkzMDE4MjIyNVoXDTMwMDkzMDE4MzIyNVowfDEL
'' SIG '' MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
'' SIG '' EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
'' SIG '' c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9z
'' SIG '' b2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggIiMA0GCSqG
'' SIG '' SIb3DQEBAQUAA4ICDwAwggIKAoICAQDk4aZM57RyIQt5
'' SIG '' osvXJHm9DtWC0/3unAcH0qlsTnXIyjVX9gF/bErg4r25
'' SIG '' PhdgM/9cT8dm95VTcVrifkpa/rg2Z4VGIwy1jRPPdzLA
'' SIG '' EBjoYH1qUoNEt6aORmsHFPPFdvWGUNzBRMhxXFExN6AK
'' SIG '' OG6N7dcP2CZTfDlhAnrEqv1yaa8dq6z2Nr41JmTamDu6
'' SIG '' GnszrYBbfowQHJ1S/rboYiXcag/PXfT+jlPP1uyFVk3v
'' SIG '' 3byNpOORj7I5LFGc6XBpDco2LXCOMcg1KL3jtIckw+DJ
'' SIG '' j361VI/c+gVVmG1oO5pGve2krnopN6zL64NF50ZuyjLV
'' SIG '' wIYwXE8s4mKyzbnijYjklqwBSru+cakXW2dg3viSkR4d
'' SIG '' Pf0gz3N9QZpGdc3EXzTdEonW/aUgfX782Z5F37ZyL9t9
'' SIG '' X4C626p+Nuw2TPYrbqgSUei/BQOj0XOmTTd0lBw0gg/w
'' SIG '' EPK3Rxjtp+iZfD9M269ewvPV2HM9Q07BMzlMjgK8Qmgu
'' SIG '' EOqEUUbi0b1qGFphAXPKZ6Je1yh2AuIzGHLXpyDwwvoS
'' SIG '' CtdjbwzJNmSLW6CmgyFdXzB0kZSU2LlQ+QuJYfM2BjUY
'' SIG '' hEfb3BvR/bLUHMVr9lxSUV0S2yW6r1AFemzFER1y7435
'' SIG '' UsSFF5PAPBXbGjfHCBUYP3irRbb1Hode2o+eFnJpxq57
'' SIG '' t7c+auIurQIDAQABo4IB3TCCAdkwEgYJKwYBBAGCNxUB
'' SIG '' BAUCAwEAATAjBgkrBgEEAYI3FQIEFgQUKqdS/mTEmr6C
'' SIG '' kTxGNSnPEP8vBO4wHQYDVR0OBBYEFJ+nFV0AXmJdg/Tl
'' SIG '' 0mWnG1M1GelyMFwGA1UdIARVMFMwUQYMKwYBBAGCN0yD
'' SIG '' fQEBMEEwPwYIKwYBBQUHAgEWM2h0dHA6Ly93d3cubWlj
'' SIG '' cm9zb2Z0LmNvbS9wa2lvcHMvRG9jcy9SZXBvc2l0b3J5
'' SIG '' Lmh0bTATBgNVHSUEDDAKBggrBgEFBQcDCDAZBgkrBgEE
'' SIG '' AYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYw
'' SIG '' DwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbL
'' SIG '' j+iiXGJo0T2UkFvXzpoYxDBWBgNVHR8ETzBNMEugSaBH
'' SIG '' hkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
'' SIG '' bC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0y
'' SIG '' My5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUFBzAC
'' SIG '' hj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
'' SIG '' cnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDAN
'' SIG '' BgkqhkiG9w0BAQsFAAOCAgEAnVV9/Cqt4SwfZwExJFvh
'' SIG '' nnJL/Klv6lwUtj5OR2R4sQaTlz0xM7U518JxNj/aZGx8
'' SIG '' 0HU5bbsPMeTCj/ts0aGUGCLu6WZnOlNN3Zi6th542DYu
'' SIG '' nKmCVgADsAW+iehp4LoJ7nvfam++Kctu2D9IdQHZGN5t
'' SIG '' ggz1bSNU5HhTdSRXud2f8449xvNo32X2pFaq95W2KFUn
'' SIG '' 0CS9QKC/GbYSEhFdPSfgQJY4rPf5KYnDvBewVIVCs/wM
'' SIG '' nosZiefwC2qBwoEZQhlSdYo2wh3DYXMuLGt7bj8sCXgU
'' SIG '' 6ZGyqVvfSaN0DLzskYDSPeZKPmY7T7uG+jIa2Zb0j/aR
'' SIG '' AfbOxnT99kxybxCrdTDFNLB62FD+CljdQDzHVG2dY3RI
'' SIG '' LLFORy3BFARxv2T5JL5zbcqOCb2zAVdJVGTZc9d/HltE
'' SIG '' AY5aGZFrDZ+kKNxnGSgkujhLmm77IVRrakURR6nxt67I
'' SIG '' 6IleT53S0Ex2tVdUCbFpAUR+fKFhbHP+CrvsQWY9af3L
'' SIG '' wUFJfn6Tvsv4O+S3Fb+0zj6lMVGEvL8CwYKiexcdFYmN
'' SIG '' cP7ntdAoGokLjzbaukz5m/8K6TT4JDVnK+ANuOaMmdbh
'' SIG '' IurwJ0I9JZTmdHRbatGePu1+oDEzfbzL6Xu/OHBE0ZDx
'' SIG '' yKs6ijoIYn/ZcGNTTY3ugm2lBRDBcQZqELQdVTNYs6Fw
'' SIG '' ZvKhggLOMIICNwIBATCB+KGB0KSBzTCByjELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFt
'' SIG '' ZXJpY2EgT3BlcmF0aW9uczEmMCQGA1UECxMdVGhhbGVz
'' SIG '' IFRTUyBFU046RTVBNi1FMjdDLTU5MkUxJTAjBgNVBAMT
'' SIG '' HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoB
'' SIG '' ATAHBgUrDgMCGgMVANGPgsi3sxoFR1hTZiiNS7hP4WOr
'' SIG '' oIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
'' SIG '' Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
'' SIG '' BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
'' SIG '' A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
'' SIG '' MTAwDQYJKoZIhvcNAQEFBQACBQDmfMdFMCIYDzIwMjIw
'' SIG '' NzE2MTMxMjM3WhgPMjAyMjA3MTcxMzEyMzdaMHcwPQYK
'' SIG '' KwYBBAGEWQoEATEvMC0wCgIFAOZ8x0UCAQAwCgIBAAIC
'' SIG '' GVUCAf8wBwIBAAICEfQwCgIFAOZ+GMUCAQAwNgYKKwYB
'' SIG '' BAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQAC
'' SIG '' AwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0BAQUFAAOB
'' SIG '' gQC25bfOa86qufLdvn3Qe5k7IlG2b3mjl48jtVsymupB
'' SIG '' 2pv7atZE5teE/pgicVHvd9jHMLgc9AmycajolTOUwHKZ
'' SIG '' N1rTN60g/XnNHkznx8CpUjGfpyQTIjF0FlUQN7S9SEJh
'' SIG '' UC3MHbtUb30mrohAmuAbBY2oJuErjRbJlbq76UpUajGC
'' SIG '' BA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwAhMzAAABlbf8DdbjNzElAAEAAAGVMA0GCWCG
'' SIG '' SAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZI
'' SIG '' hvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIC9gWe1BiUJ9
'' SIG '' 0FJtONt0aUZoVTN6M3ocj9We90HV17teMIH6BgsqhkiG
'' SIG '' 9w0BCRACLzGB6jCB5zCB5DCBvQQgXOZL4Y2QC3tpoSM/
'' SIG '' 0He5HlTpgP3AtXcymU+MmyxJAscwgZgwgYCkfjB8MQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3Nv
'' SIG '' ZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAZW3/A3W
'' SIG '' 4zcxJQABAAABlTAiBCDgbeBRXgMMtzQWQMXQ5+5btGYH
'' SIG '' 8cE4neSTkgSApjaEqjANBgkqhkiG9w0BAQsFAASCAgBE
'' SIG '' YTkrZ01pF5dD3z1tpMGFuRr1oOfeIF6kipha6Whpny6p
'' SIG '' B0a6iW7fPW3WTztBA24h9Hrz/TM1WxFnvcQnTir80rQg
'' SIG '' J67e7TFjQljag9I94BSx5vYO254BOuO9gwKGQkwjxDUE
'' SIG '' y7MsGcJbt10C+qdOsOtLsUXc1f7KkBPGEFCYTzJHC1ly
'' SIG '' D8pJOb33zuli28xk9MfkoqzblzAj4ScrPTCYLMDBq2gf
'' SIG '' TTJGipgRVB55b5EbPF7qLi4FZiQ0HI19rerreLomfp1U
'' SIG '' k5P5vjx8Bg8X9K51deZYE3nPr4aX3zDD2qlM74tj0win
'' SIG '' dk2m+DZw653sbwxLfvHTQX83ocaKcjtvPAmNrWvu4Jqh
'' SIG '' R+eBMYji4pspAWDp1dsqbGHfl6KbjJsjE5MNBnJ99oAE
'' SIG '' bEl56kJyoxjH7axXUMcX133bPbWBbxcQh3ppAzoATA0O
'' SIG '' DotJTbo4Lp+B47czuKAzAa+B2b03TC7aO6U53uTyzBQY
'' SIG '' jAVWbYokcNtT4oslfYo8ZdwRm+XbAuGXiqggc2b6mMv+
'' SIG '' 0n1YidoNWoDsmbgoQqvRcGA+e4WH5m2HECX0M8o6YY3U
'' SIG '' +165W/7sEGq28ohZEZEYZwt4dGWm6kofQWwM9f/f2rsm
'' SIG '' qGv3w+iD/m7cdbab6Uqflj506PyeQ+nZboOzHeaJwPV5
'' SIG '' 9T+XqEjN3VWTcS4VWSysZg==
'' SIG '' End signature block

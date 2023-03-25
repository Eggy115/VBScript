' Windows Installer utility to preview dialogs from a install database
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of preview APIs
'
Option Explicit

Const msiOpenDatabaseModeReadOnly = 0

' Show help if no arguments or if argument contains ?
Dim argCount : argCount = Wscript.Arguments.Count
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If argCount = 0 Then
	Wscript.Echo "Windows Installer utility to preview dialogs from an install database." &_
		vbLf & " The 1st argument is the path to an install database, relative or complete path" &_
		vbLf & " Subsequent arguments are dialogs to display (primary key of Dialog table)" &_
		vbLf & " To show a billboard, append the Control name (Control table key) and Billboard" &_
		vbLf & "       name (Billboard table key) to the Dialog name, separated with colons." &_
		vbLf & " If no dialogs specified, all dialogs in Dialog table are displayed sequentially" &_
		vbLf & " Note: The name of the dialog, if provided,  is case-sensitive" &_
		vblf &_
		vblf & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

' Open database
Dim databasePath : databasePath = Wscript.Arguments(0)
Dim database : Set database = installer.OpenDatabase(databasePath, msiOpenDatabaseModeReadOnly) : CheckError

' Create preview object
Dim preview : Set preview = Database.EnableUIpreview : CheckError

' Get properties from Property table and put into preview object
Dim record, view : Set view = database.OpenView("SELECT `Property`,`Value` FROM `Property`") : CheckError
view.Execute : CheckError
Do
	Set record = view.Fetch : CheckError
	If record Is Nothing Then Exit Do
	preview.Property(record.StringData(1)) = record.StringData(2) : CheckError
Loop

' Loop through list of dialog names and display each one
If argCount = 1 Then ' No dialog name, loop through all dialogs
	Set view = database.OpenView("SELECT `Dialog` FROM `Dialog`") : CheckError
	view.Execute : CheckError
	Do
		Set record = view.Fetch : CheckError
		If record Is Nothing Then Exit Do
		preview.ViewDialog(record.StringData(1)) : CheckError
		Wait
	Loop
Else ' explicit dialog names supplied
	Set view = database.OpenView("SELECT `Dialog` FROM `Dialog` WHERE `Dialog`=?") : CheckError
	Dim paramRecord, argNum, argArray, dialogName, controlName, billboardName
	Set paramRecord = installer.CreateRecord(1)
	For argNum = 1 To argCount-1
		dialogName = Wscript.Arguments(argNum)
		argArray = Split(dialogName,":",-1,vbTextCompare)
		If UBound(argArray) <> 0 Then  ' billboard to add to dialog
			If UBound(argArray) <> 2 Then Fail "Incorrect billboard syntax, must specify 3 values"
			dialogName    = argArray(0)
			controlName   = argArray(1) ' we could validate that controlName is in the Control table
			billboardName = argArray(2) ' we could validate that billboard is in the Billboard table
		End If
		paramRecord.StringData(1) = dialogName
		view.Execute paramRecord : CheckError
		If view.Fetch Is Nothing Then Fail "Dialog not found: " & dialogName
		preview.ViewDialog(dialogName) : CheckError
		If UBound(argArray) = 2 Then preview.ViewBillboard controlName, billboardName : CheckError
		Wait
	Next
End If
preview.ViewDialog ""  ' clear dialog, must do this to release object deadlock

' Wait until user input to clear dialog. Too bad there's no function to wait for keyboard input
Sub Wait
	Dim shell : Set shell = Wscript.CreateObject("Wscript.Shell")
	MsgBox "Next",0,"Drag me away"
End Sub

Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbLf & errRec.FormatText
	End If
	Fail message
End Sub

Sub Fail(message)
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIIl8gYJKoZIhvcNAQcCoIIl4zCCJd8CAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' mOr7DzNLA7B3kQygPHKkFo0lJ4ImjipM2G/ZKh4w1cKg
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
'' SIG '' DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBA
'' SIG '' QNzG0p24TUN1L8ZDRKn3RWNowySEKvCa5C9u9njgujA8
'' SIG '' BgorBgEEAYI3CgMcMS4MLExveU1LZ3RxQm96bHUxTlFS
'' SIG '' K00xRXQ1aXZ2ZXIzSTg1M1B4b1VvM2VyM0k9MFoGCisG
'' SIG '' AQQBgjcCAQwxTDBKoCSAIgBNAGkAYwByAG8AcwBvAGYA
'' SIG '' dAAgAFcAaQBuAGQAbwB3AHOhIoAgaHR0cDovL3d3dy5t
'' SIG '' aWNyb3NvZnQuY29tL3dpbmRvd3MwDQYJKoZIhvcNAQEB
'' SIG '' BQAEggEAGwuBxUsczanyuz4QWYtr02QwBY5GHaKUEVd1
'' SIG '' rA8cOrz4QOmL8+XLPtlMFVOiRlcktHsxeHI/2GbkcN0A
'' SIG '' rodGF/zcsr5bFefCAoPlXeUjuZVlwbOpnSwObDHypvKZ
'' SIG '' N5Aafy3nqsjmwXRzCoQaU9R+tUU48dWqqSIeFqVheR3Q
'' SIG '' 9mQ2ypxAaTEe/vO0uGHqjmyy+SYj2U/C/5exo3+jvVaT
'' SIG '' gFi5QKY+0+hj8UU6cLjTPt33LYjkUFzsmRM0ZXSMe6Pv
'' SIG '' ynbWw8o+qdXzZk2oQk4pjOA/9Tz5btkpk6rakU7gEh+p
'' SIG '' fTtJwxehL88o0B5OZMekPVOv7cuG6Il2u+Lxq6A++6GC
'' SIG '' FwAwghb8BgorBgEEAYI3AwMBMYIW7DCCFugGCSqGSIb3
'' SIG '' DQEHAqCCFtkwghbVAgEDMQ8wDQYJYIZIAWUDBAIBBQAw
'' SIG '' ggFRBgsqhkiG9w0BCRABBKCCAUAEggE8MIIBOAIBAQYK
'' SIG '' KwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCBkiCcP
'' SIG '' Tw9HATKpnMtULvbdMxV4lhaEbAin2QIamc1lxQIGYrSf
'' SIG '' u9/wGBMyMDIyMDcxNjA4NTY1Ny45NTNaMASAAgH0oIHQ
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
'' SIG '' hvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIDGZlCLwjNAJ
'' SIG '' Oc8nrjVIqGDQwON0DN4Mith+0s+XPHZQMIH6BgsqhkiG
'' SIG '' 9w0BCRACLzGB6jCB5zCB5DCBvQQgXOZL4Y2QC3tpoSM/
'' SIG '' 0He5HlTpgP3AtXcymU+MmyxJAscwgZgwgYCkfjB8MQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3Nv
'' SIG '' ZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAZW3/A3W
'' SIG '' 4zcxJQABAAABlTAiBCDgbeBRXgMMtzQWQMXQ5+5btGYH
'' SIG '' 8cE4neSTkgSApjaEqjANBgkqhkiG9w0BAQsFAASCAgAQ
'' SIG '' niGjpcNUlg32wsV0RpCn6rS9UdGwkkyhF8TewbuKfHCc
'' SIG '' zyQbgkbMXMJqSRDAgKpHSvvrzRATWAurRAUa30yHZfmV
'' SIG '' ClGGUoeIzx/KSHZe7NzWOcw7ksMY9+uQqx16jPFHCq97
'' SIG '' SMQ1qNbO5Jk8VDbw6q+P6P7bCOZ6dBWIvmLx9MWtSqvl
'' SIG '' KnOkyftSCTgffOuCoxmWqEsJOtMiFJ05O3GsEi3mMTsb
'' SIG '' q+yS+iUfPbC7bcbYuuSuFipqrTOalMd6tcmea00x/LpI
'' SIG '' PDAspgR7iCkqtnJypbtt4bzo7hiymDupyjY4M82+IzRn
'' SIG '' 0oGtqHx//tABWTPUJ/hgNkJa8BTvyA27ROBqVC885wCH
'' SIG '' VbZWqc/xIzPT/tEZcyzZARVEgD7qMSvQsC9IKlQ2cWvW
'' SIG '' i0W0FU/3y5KtLr4wG5JesOR0ueH8Wu8c4fV3J5Ni7KjW
'' SIG '' L0ZQ/SahKhIhWFSiX1WCnUPgRjxUPbc2Qr8PZF4HSvb4
'' SIG '' 2kMcLxyPDUciPaWX0xvLvm2k675cW/z2hc3K1aZRBH0z
'' SIG '' D7TX3sf4P6EBnc+uAb86Zg/DgeRCWz5lGAQa82UOuOJB
'' SIG '' aaj38c3bt7/kiLn7zbtwFGzapwwFBguQyEVcLAytDafH
'' SIG '' OPz3Q0jbDkjA7mCMZqToIgOSk1XpK+kK4eSSekG62MKy
'' SIG '' wpGj4nXZR+RSaSVdCU8tQQ==
'' SIG '' End signature block

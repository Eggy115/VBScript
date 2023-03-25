' Windows Installer utility to report the differences between two databases
' For use with Windows Scripting Host, CScript.exe only, lists to stdout
' Copyright (c) Microsoft Corporation. All rights reserved.
' Simply generates a transform between the databases and then view the transform
'
Option Explicit

Const icdLong       = 0
Const icdShort      = &h400
Const icdObject     = &h800
Const icdString     = &hC00
Const icdNullable   = &h1000
Const icdPrimaryKey = &h2000
Const icdNoNulls    = &h0000
Const icdPersistent = &h0100
Const icdTemporary  = &h0000

Const msiOpenDatabaseModeReadOnly     = 0
Const msiOpenDatabaseModeTransact     = 1
Const msiOpenDatabaseModeCreate       = 3
Const iteViewTransform       = 256

If Wscript.Arguments.Count < 2 Then
	Wscript.Echo "Windows Installer database difference utility" &_
		vbNewLine & " Generates a temporary transform file, then display it" &_
		vbNewLine & " 1st argument is the path to the original installer database" &_
		vbNewLine & " 2nd argument is the path to the updated installer database" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Cannot run with GUI script host, as listing is performed to standard out
If UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "W" Then
	WScript.Echo "Cannot use WScript.exe - must use CScript.exe with this program"
	Wscript.Quit 2
End If

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

' Create path for temporary transform file
Dim WshShell : Set WshShell = Wscript.CreateObject("Wscript.Shell") : CheckError
Dim tempFilePath:tempFilePath = WshShell.ExpandEnvironmentStrings("%TEMP%") & "\diff.tmp"

' Open databases, generate transform, then list transform
Dim database1 : Set database1 = installer.OpenDatabase(Wscript.Arguments(0), msiOpenDatabaseModeReadOnly) : CheckError
Dim database2 : Set database2 = installer.OpenDatabase(Wscript.Arguments(1), msiOpenDatabaseModeReadOnly) : CheckError
Dim different : different = Database2.GenerateTransform(Database1, tempFilePath) : CheckError
If different Then
	database1.ApplyTransform tempFilePath, iteViewTransform + 0 : CheckError' should not need error suppression flags
	ListTransform database1
End If

' Open summary information streams and compare them
Dim sumInfo1 : Set sumInfo1 = database1.SummaryInformation(0) : CheckError
Dim sumInfo2 : Set sumInfo2 = database2.SummaryInformation(0) : CheckError
Dim iProp, value1, value2
For iProp = 1 to 19              
	value1 = sumInfo1.Property(iProp) : CheckError
	value2 = sumInfo2.Property(iProp) : CheckError
	If value1 <> value2 Then
		Wscript.Echo "\005SummaryInformation   [" & iProp & "] {" & value1 & "}->{" & value2 & "}"
		different = True
	End If
Next
If Not different Then Wscript.Echo "Databases are identical"
Wscript.Quit 0

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
	Set view = database.OpenView("SELECT * FROM `_TransformView` ORDER BY `Table`, `Row`")
	If Err <> 0 Then Wscript.Echo "Transform viewing supported only in builds 4906 and beyond of MSI.DLL" : Wscript.Quit 2
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
'' SIG '' MIIl8gYJKoZIhvcNAQcCoIIl4zCCJd8CAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' cwzPLmwXndAasgrSds09lWoEa+nEByy+weD1dY9VbGig
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
'' SIG '' DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCAc
'' SIG '' JM84WRJBZUuta8N3M7qSH3vekAlmANl0yXCJB8IWdjA8
'' SIG '' BgorBgEEAYI3CgMcMS4MLG5uQzlHWDZucDVpdlc4a3lJ
'' SIG '' Zm95NGdEbmNYM0dyd2k0N2V0WmtWWUtsVXc9MFoGCisG
'' SIG '' AQQBgjcCAQwxTDBKoCSAIgBNAGkAYwByAG8AcwBvAGYA
'' SIG '' dAAgAFcAaQBuAGQAbwB3AHOhIoAgaHR0cDovL3d3dy5t
'' SIG '' aWNyb3NvZnQuY29tL3dpbmRvd3MwDQYJKoZIhvcNAQEB
'' SIG '' BQAEggEAazqpm1oxW/D+imyLZBuo2Tq8FRzKtFW0PRnv
'' SIG '' Fv6QPM/KIaM2mZiWL+2l04C0tQzO8cL2NHeub0o/A7GD
'' SIG '' WKP6Lgx5hFm+2kRSpSYaYHP6hIm5by1+wrDibHMEAxqa
'' SIG '' AFbfAfUPZvINxSbgTU2ZebHGfXR/0JKA79tcGZJ/OPDq
'' SIG '' 0bAydtcDTlXvo4amDOQC2iFU7YZC3jKJXy0XH75cAO+Z
'' SIG '' VMlxf8W/jgZ9+x9Yt9ZTVG61o4edb2F4djLw6R5IcZyN
'' SIG '' YU42iuga68zK03fWuoUfC8OuH+vdzueS3y6VHsPK5Skf
'' SIG '' WS0kHy7QzXmSbG63P/Gk8wthKJ5zSOF7ZXovmTlkaKGC
'' SIG '' FwAwghb8BgorBgEEAYI3AwMBMYIW7DCCFugGCSqGSIb3
'' SIG '' DQEHAqCCFtkwghbVAgEDMQ8wDQYJYIZIAWUDBAIBBQAw
'' SIG '' ggFRBgsqhkiG9w0BCRABBKCCAUAEggE8MIIBOAIBAQYK
'' SIG '' KwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCDJxdb2
'' SIG '' 9usPABS585d09K2ur4xFJqsAOcVnra1mZA4fCwIGYrSI
'' SIG '' jKldGBMyMDIyMDcxNjA4NTY1OC4zMDJaMASAAgH0oIHQ
'' SIG '' pIHNMIHKMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
'' SIG '' aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
'' SIG '' ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
'' SIG '' ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMSYw
'' SIG '' JAYDVQQLEx1UaGFsZXMgVFNTIEVTTjoxMkJDLUUzQUUt
'' SIG '' NzRFQjElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3Rh
'' SIG '' bXAgU2VydmljZaCCEVcwggcMMIIE9KADAgECAhMzAAAB
'' SIG '' oQGFVZm5VF2KAAEAAAGhMA0GCSqGSIb3DQEBCwUAMHwx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
'' SIG '' c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTIxMTIw
'' SIG '' MjE5MDUyNFoXDTIzMDIyODE5MDUyNFowgcoxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
'' SIG '' bWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxl
'' SIG '' cyBUU1MgRVNOOjEyQkMtRTNBRS03NEVCMSUwIwYDVQQD
'' SIG '' ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIIC
'' SIG '' IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA2sk8
'' SIG '' XuVrpJK2McVbhy2FvQRFgg/ZJI55x7DisBnXSD22ZS2P
'' SIG '' peaLywzX/gRDECgGUCNw1/dZdcgg7j/V+7TjwuPGURlw
'' SIG '' P23/apdBSueN/ICJe3FedvF3hDhcHPwPlGyFH1tvejpo
'' SIG '' PGetsWkL946xuFP6a4gKxf3q9VANRzbiBlMqo5coIkj8
'' SIG '' CtjZxQKYtSQ/lHn+XOO5Ie6VtSo+0Z3IaRXmPTHpD0EY
'' SIG '' mu3BGlGFOLKgoiVXQyaXny7z0/RHbYZUMe+ZXcfgMGX9
'' SIG '' mvU+7kEUgYfLacT3SAw5ColjMIyk6wGNPQNyP44naj7n
'' SIG '' PD71/rKsasmRDdoeBgNBHY5pOuJ5CLpACtfCuZwCwyzv
'' SIG '' UjE8aQMECB0Q7WXkwpbwDwhKMtb7Tw+3/nqh6krbrvlw
'' SIG '' pH0Y1xKV/fofX67AdPwYA+QgX9xCywGvE3nzHx2VhCUU
'' SIG '' zza21zCos0q1EpFb/9xz/2bCacGs+TMtkW8nNwIfW0++
'' SIG '' ngSZMn0+RTfb/ykNB58YUTLOhx4U5jcfi87WHIvrx39A
'' SIG '' 90B9Xgo2VmUY6dZjssaT1NpgzBuoHpbybHtSc0QA6O2C
'' SIG '' KJPydwnG5vDGwW5vOYqIBZbRR3nBxRBcK7AxgRZzWBzI
'' SIG '' XG2q0DQPoGNntpfXwJF9zIyO1JJZKM++Pz+iiKnuY3Hf
'' SIG '' RTwm20m2B/Ti7LXnmDkCAwEAAaOCATYwggEyMB0GA1Ud
'' SIG '' DgQWBBQWvyAy22OO+VUMiomUsOO5dP3MqTAfBgNVHSME
'' SIG '' GDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8E
'' SIG '' WDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
'' SIG '' b20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0
'' SIG '' YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUH
'' SIG '' AQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1p
'' SIG '' Y3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29m
'' SIG '' dCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
'' SIG '' dDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsGAQUF
'' SIG '' BwMIMA0GCSqGSIb3DQEBCwUAA4ICAQAgwyMRJX15RuWC
'' SIG '' nCqyz0kvn9yVmaa8ChIgEr4r2h0wUZV7QLPk5GnXLBHo
'' SIG '' vvcOb5hebQlM0x+HNJwiO22cZ7C/kul7IjrN2dVeFl/i
'' SIG '' MKF1CeMy77NPpk+L4xg7WHykP27JiSmq9nPfZv3x79Vp
'' SIG '' tgk3Mmnj74vOiYd1Mi43USC1m7c7OKCJhTMMCm8x3T6K
'' SIG '' cawYYIvgtWGbIaLFi5YM8rsY1JfqjYNZudjCZn9dZaCO
'' SIG '' w/RyaGkM3fq3/dvGPK71C5oNofxudKPg9FCdRWv3CSWh
'' SIG '' 3wd7HysPV+hq7V2Bo5jN/oPgIWlbH7qSlzbThbubZyyr
'' SIG '' wB+TiIxA2FdWCppV7gboW2GrLMoDxTJjYBtgJ5N3axHA
'' SIG '' 3GYQl16qUbMzaNRehruSQqUGV2ziTPVHuT5SSrZiJgGC
'' SIG '' BrMPqZx8v6+YIEmDqeIOWdaFPRoVQjN1dE/WnXnujlFw
'' SIG '' ZNaxOHWXP1LD5Y9KqIpYy/pTdQOYJJps+5ObSDm1Rge3
'' SIG '' SXc/CdBcF0ROamLtQHb2rlW2cBkJC9cfGiv7L4xEFtDV
'' SIG '' Midvc5wx4l5eby6EU44xabIVAYtviGPpjamy5o9uI+Xk
'' SIG '' /m4w5RNx5jbSz6S3DA2KmdR/ulOmJmojZmnNo0VwwGnh
'' SIG '' BP7qAzLdnQK3yT+zPjA7988zTUyDXrjRLQ1YJvc8H4CF
'' SIG '' Al5w2blbYjCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
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
'' SIG '' IFRTUyBFU046MTJCQy1FM0FFLTc0RUIxJTAjBgNVBAMT
'' SIG '' HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoB
'' SIG '' ATAHBgUrDgMCGgMVABtxdozuCxDFS8IChl3WDDeBQYDg
'' SIG '' oIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
'' SIG '' Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
'' SIG '' BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
'' SIG '' A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
'' SIG '' MTAwDQYJKoZIhvcNAQEFBQACBQDmfLAIMCIYDzIwMjIw
'' SIG '' NzE2MTEzMzI4WhgPMjAyMjA3MTcxMTMzMjhaMHcwPQYK
'' SIG '' KwYBBAGEWQoEATEvMC0wCgIFAOZ8sAgCAQAwCgIBAAIC
'' SIG '' GkkCAf8wBwIBAAICEi4wCgIFAOZ+AYgCAQAwNgYKKwYB
'' SIG '' BAGEWQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQAC
'' SIG '' AwehIKEKMAgCAQACAwGGoDANBgkqhkiG9w0BAQUFAAOB
'' SIG '' gQBYJjx9tMH0TNSNeSeLGGmIW/b/4YHVT6O43jf3lXky
'' SIG '' FUXDPaBYO/aF0ExZjfUk3yetCL4WtlyF1Y8Ger7o8bF1
'' SIG '' Bgkz4e451aT56lJD8DZoTA1Z5XlVLTTuwcbZoY6aRw/V
'' SIG '' FBrsdq4ahFzS1dP6MUloGdtWxx9QCp9IgUHmJU979DGC
'' SIG '' BA0wggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYD
'' SIG '' VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
'' SIG '' MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24x
'' SIG '' JjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBD
'' SIG '' QSAyMDEwAhMzAAABoQGFVZm5VF2KAAEAAAGhMA0GCWCG
'' SIG '' SAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZI
'' SIG '' hvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIAk+3YrtMxLq
'' SIG '' GMds0hKj5NPmyBq+h5yEm7FDNtJB5a1mMIH6BgsqhkiG
'' SIG '' 9w0BCRACLzGB6jCB5zCB5DCBvQQg6whU8TqBgmgggo6E
'' SIG '' cgXtSUkKzCXggk8hK84oid+O0IQwgZgwgYCkfjB8MQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3Nv
'' SIG '' ZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAaEBhVWZ
'' SIG '' uVRdigABAAABoTAiBCAGXZ59hdVCRxj+jjABiXIy2kc5
'' SIG '' dFDc8b1nwMqVAIx7zDANBgkqhkiG9w0BAQsFAASCAgCd
'' SIG '' yt0ClMAoRSA++vUdryTN/U6VNZMoTzc5dtHUbEZUICvH
'' SIG '' D7K5H1+CDM4vx29744KnrGaMuZHzSw+CJqxXuDRwc2sO
'' SIG '' QX19nVBGMFAjnMYrvopXeCBWiinW3V7sxt6spFKNvnpV
'' SIG '' NPKrqbtnheu9+/Nmiyn8j/8H9IyTAO6UZBWzEmsNZccc
'' SIG '' +3+kFSMqDQvGOP3tMjrYgQcgik/OvjEpwEYNc+2uFMOG
'' SIG '' 1rxGltJ431zXMbxnjKgbIFEG5M+WFCckRcPn6ooUQkK9
'' SIG '' xNL03pOgBm8yeT8YulrZcKV2uy3R3OlORJYKjQQKdBai
'' SIG '' C1hwYfCkMy0WwVqNGqsyW/Tf3ndY9rPF96UbUMZN4uwm
'' SIG '' J+Z1D86/EcsTUWGX2NvLivD/h74mAnL9t5s8Nd8blzlA
'' SIG '' /oO1YHw/GjjOoYXCx252YfT3Pc5q3EqkHHKy2/n8OzIX
'' SIG '' ZgXTrEIm1LIjjTzfExVMn3dpndB7WjkU1/oULZhZ67vr
'' SIG '' oLmPAX9xL47ntg/kNWo9ym+hTOv+OKQ3ribHp/dDXjug
'' SIG '' NwxMYA9Y0HSpDxfx9fEhVNr8ZqgAUNrrMmtShu73rMD4
'' SIG '' 7UTYDGJkC+VKNG0liHsfnbVwVsLrEMgyzVaf7XRxbMDl
'' SIG '' jhN13OPEHwyCohzUXZGqZy0RheEHTX+hVTrZCI8yoFD7
'' SIG '' aJ7uOvQDnbAfa/CoRjiueg==
'' SIG '' End signature block

' Windows Installer utility to list registered products and product info
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the use of the product enumeration and ProductInfo methods and underlying APIs
'
Option Explicit

Const msiInstallStateNotUsed      = -7
Const msiInstallStateBadConfig    = -6
Const msiInstallStateIncomplete   = -5
Const msiInstallStateSourceAbsent = -4
Const msiInstallStateInvalidArg   = -2
Const msiInstallStateUnknown      = -1
Const msiInstallStateBroken       =  0
Const msiInstallStateAdvertised   =  1
Const msiInstallStateRemoved      =  1
Const msiInstallStateAbsent       =  2
Const msiInstallStateLocal        =  3
Const msiInstallStateSource       =  4
Const msiInstallStateDefault      =  5

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

' If no arguments supplied, then list all installed or advertised products
Dim argCount:argCount = Wscript.Arguments.Count
If (argCount = 0) Then
	Dim product, products, info, productList, version
	On Error Resume Next
	Set products = installer.Products : CheckError
	For Each product In products
		version = DecodeVersion(installer.ProductInfo(product, "Version")) : CheckError
		info = product & " = " & installer.ProductInfo(product, "ProductName") & " " & version : CheckError
		If productList <> Empty Then productList = productList & vbNewLine & info Else productList = info
	Next
	If productList = Empty Then productList = "No products installed or advertised"
	Wscript.Echo productList
	Set products = Nothing
	Wscript.Quit 0
End If

' Check for ?, and show help message if found
Dim productName:productName = Wscript.Arguments(0)
If InStr(1, productName, "?", vbTextCompare) > 0 Then
	Wscript.Echo "Windows Installer utility to list registered products and product information" &_
		vbNewLine & " Lists all installed and advertised products if no arguments are specified" &_
		vbNewLine & " Else 1st argument is a product name (case-insensitive) or product ID (GUID)" &_
		vbNewLine & " If 2nd argument is missing or contains 'p', then product properties are listed" &_
		vbNewLine & " If 2nd argument contains 'f', features, parents, & installed states are listed" &_
		vbNewLine & " If 2nd argument contains 'c', installed components for that product are listed" &_
		vbNewLine & " If 2nd argument contains 'd', HKLM ""SharedDlls"" count for key files are listed" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' If Product name supplied, need to search for product code
Dim productCode, property, value, message
If Left(productName, 1) = "{" And Right(productName, 1) = "}" Then
	If installer.ProductState(productName) <> msiInstallStateUnknown Then productCode = UCase(productName)
Else
	For Each productCode In installer.Products : CheckError
		If LCase(installer.ProductInfo(productCode, "ProductName")) = LCase(productName) Then Exit For
	Next
End If
If IsEmpty(productCode) Then Wscript.Echo "Product is not registered: " & productName : Wscript.Quit 2

' Check option argument for type of information to display, default is properties
Dim optionFlag : If argcount > 1 Then optionFlag = LCase(Wscript.Arguments(1)) Else optionFlag = "p"
If InStr(1, optionFlag, "*", vbTextCompare) > 0 Then optionFlag = "pfcd"

If InStr(1, optionFlag, "p", vbTextCompare) > 0 Then
	message = "ProductCode = " & productCode
	For Each property In Array(_
			"Language",_
			"ProductName",_
			"PackageCode",_
			"Transforms",_
			"AssignmentType",_
			"PackageName",_
			"InstalledProductName",_
			"VersionString",_
			"RegCompany",_
			"RegOwner",_
			"ProductID",_
			"ProductIcon",_
			"InstallLocation",_
			"InstallSource",_
			"InstallDate",_
			"Publisher",_
			"LocalPackage",_
			"HelpLink",_
			"HelpTelephone",_
			"URLInfoAbout",_
			"URLUpdateInfo") : CheckError
		value = installer.ProductInfo(productCode, property) ': CheckError
		If Err <> 0 Then Err.Clear : value = Empty
		If (property = "Version") Then value = DecodeVersion(value)
		If value <> Empty Then message = message & vbNewLine & property & " = " & value
	Next
	Wscript.Echo message
End If

If InStr(1, optionFlag, "f", vbTextCompare) > 0 Then
	Dim feature, features, parent, state, featureInfo
	Set features = installer.Features(productCode)
	message = "---Features in product " & productCode & "---"
	For Each feature In features
		parent = installer.FeatureParent(productCode, feature) : CheckError
		If Len(parent) Then parent = " {" & parent & "}"
		state = installer.FeatureState(productCode, feature)
		Select Case(state)
			Case msiInstallStateBadConfig:    state = "Corrupt"
			Case msiInstallStateIncomplete:   state = "InProgress"
			Case msiInstallStateSourceAbsent: state = "SourceAbsent"
			Case msiInstallStateBroken:       state = "Broken"
			Case msiInstallStateAdvertised:   state = "Advertised"
			Case msiInstallStateAbsent:       state = "Uninstalled"
			Case msiInstallStateLocal:        state = "Local"
			Case msiInstallStateSource:       state = "Source"
			Case msiInstallStateDefault:      state = "Default"
			Case Else:                        state = "Unknown"
		End Select
		message = message & vbNewLine & feature & parent & " = " & state
	Next
	Set features = Nothing
	Wscript.Echo message
End If 

If InStr(1, optionFlag, "c", vbTextCompare) > 0 Then
	Dim component, components, client, clients, path
	Set components = installer.Components : CheckError
	message = "---Components in product " & productCode & "---"
	For Each component In components
		Set clients = installer.ComponentClients(component) : CheckError
		For Each client In Clients
			If client = productCode Then
				path = installer.ComponentPath(productCode, component) : CheckError
				message = message & vbNewLine & component & " = " & path
				Exit For
			End If
		Next
		Set clients = Nothing
	Next
	Set components = Nothing
	Wscript.Echo message
End If

If InStr(1, optionFlag, "d", vbTextCompare) > 0 Then
	Set components = installer.Components : CheckError
	message = "---Shared DLL counts for key files of " & productCode & "---"
	For Each component In components
		Set clients = installer.ComponentClients(component) : CheckError
		For Each client In Clients
			If client = productCode Then
				path = installer.ComponentPath(productCode, component) : CheckError
				If Len(path) = 0 Then path = "0"
				If AscW(path) >= 65 Then  ' ignore registry key paths
					value = installer.RegistryValue(2, "SOFTWARE\Microsoft\Windows\CurrentVersion\SharedDlls", path)
					If Err <> 0 Then value = 0 : Err.Clear
					message = message & vbNewLine & value & " = " & path
				End If
				Exit For
			End If
		Next
		Set clients = Nothing
	Next
	Set components = Nothing
	Wscript.Echo message
End If

Function DecodeVersion(version)
	version = CLng(version)
	DecodeVersion = version\65536\256 & "." & (version\65535 MOD 256) & "." & (version Mod 65536)
End Function

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
'' SIG '' +TYCWFk7lUqBMQntWKZoHVk2tbD50YMJse1NdDP1q+Gg
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
'' SIG '' DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBX
'' SIG '' fwquaSys1mZF7Y8hGEH6M51b0VZVajJUmLrUJX98MDA8
'' SIG '' BgorBgEEAYI3CgMcMS4MLDBQRDZDVmlOWUtUWEd2NE52
'' SIG '' MnNIc0Nndmt3SThkUHRzdVRkbHo2NHpmOG89MFoGCisG
'' SIG '' AQQBgjcCAQwxTDBKoCSAIgBNAGkAYwByAG8AcwBvAGYA
'' SIG '' dAAgAFcAaQBuAGQAbwB3AHOhIoAgaHR0cDovL3d3dy5t
'' SIG '' aWNyb3NvZnQuY29tL3dpbmRvd3MwDQYJKoZIhvcNAQEB
'' SIG '' BQAEggEAQ/WkCVnaQoA64ZKmEexkydaOa0JnqCSkkc9g
'' SIG '' rN3dN4oJl9pL/o7hkN9xAo4JofCSQicc+l+9d4WShPGO
'' SIG '' BblAZbKf5kuUdqVmFH0QycbmKdNes1cGwA1cP0UJAANI
'' SIG '' PzjYZYx9ayfah0mNA5K5GSsV6CMkpMbO0RKdwQa8IzuP
'' SIG '' 9By1q3c18A66R6lwx4evrw3l+456OZ5IKA0Wa3MhvspJ
'' SIG '' HYIpqlW6LPp8ItFEHbAbxas7mBimVzUo6vAm5rhdbUfQ
'' SIG '' ExSyCXnmoI0W7KxxCv5NmvkOgfJ7T2BaJ/Wn00ZgZj0F
'' SIG '' 3udV0nkfqXXxm8PsLk9SdwZFcmckvDTt8n34phYdxaGC
'' SIG '' FwAwghb8BgorBgEEAYI3AwMBMYIW7DCCFugGCSqGSIb3
'' SIG '' DQEHAqCCFtkwghbVAgEDMQ8wDQYJYIZIAWUDBAIBBQAw
'' SIG '' ggFRBgsqhkiG9w0BCRABBKCCAUAEggE8MIIBOAIBAQYK
'' SIG '' KwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCAIptg9
'' SIG '' NFBzsBU4S6fNGqwruWjTQpZ2sbnB1NkwmNjMpQIGYrSf
'' SIG '' u+C7GBMyMDIyMDcxNjA4NTcwMS40MzRaMASAAgH0oIHQ
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
'' SIG '' hvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIIEqVYAgYGjN
'' SIG '' rhZCfLli8H+XMTcWiaqZd7EtUwwYcPDWMIH6BgsqhkiG
'' SIG '' 9w0BCRACLzGB6jCB5zCB5DCBvQQgXOZL4Y2QC3tpoSM/
'' SIG '' 0He5HlTpgP3AtXcymU+MmyxJAscwgZgwgYCkfjB8MQsw
'' SIG '' CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
'' SIG '' MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9z
'' SIG '' b2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3Nv
'' SIG '' ZnQgVGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAZW3/A3W
'' SIG '' 4zcxJQABAAABlTAiBCDgbeBRXgMMtzQWQMXQ5+5btGYH
'' SIG '' 8cE4neSTkgSApjaEqjANBgkqhkiG9w0BAQsFAASCAgB8
'' SIG '' R7X3m8+T4p/JF6TLYZrND3PxyrYhXOq23ykIw92vdYQu
'' SIG '' INDeO+6rkoQY2Tr7NooCInJIHna8mFvf7t56aYybJv9R
'' SIG '' JdL7EeMXXw3MorE9KpuXqwiCCPGzkg8+plpQ/q9zjIG6
'' SIG '' Toop2sBeh2KNTFFwne4X7WKDS+uUY1vtzjgbGg+007bU
'' SIG '' AkPoV7SvOAig4GaPUHHbCc0csIr/gsNCTPgxPZnaEUXX
'' SIG '' SiLh/jV8OIwzH44z9rPeXxEf2q5cX9hcE2QIgdHnMxBc
'' SIG '' IdpmtH9AasfXCmpbw95fp7HWFlcC5jCGR6PmfaNAVASO
'' SIG '' z834KAUAPYIexJu0OUTgz2a6R4ZhF3JAthuugdObnSm7
'' SIG '' X2ld48t38s7993EUCoj25wzxAJ389noORCla2LFJ50Df
'' SIG '' JWZlxPp1vG2Bx34c0cozU5Y1SC4siExT2G3Ox9y8cXyw
'' SIG '' tyJGA+938+UfiRrU+fi9UXl2kFmjwXWgMD8Qz8P5sN0w
'' SIG '' F9dr9VNlza1HNV/XJxPe1WSgn0NG5QY4wEkL39X6V9oe
'' SIG '' bfKwcXj0blKTfWvJIz4rXdqrotwLw7+tE88ZguFd7NPr
'' SIG '' UCB24LwzWPdVN2DQ7SfFf1OO6LpCkN7XANkVrIMftkQN
'' SIG '' XGolChRV08n1jGzApYk+mvhrxg8lZLXGTVtMweDQfL+d
'' SIG '' wGgttHJSzjV3Hv7G/LH6vA==
'' SIG '' End signature block

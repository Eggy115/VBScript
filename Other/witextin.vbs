' Windows Installer utility to copy a file into a database text field
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates processing of primary key data
'
Option Explicit

Const msiOpenDatabaseModeReadOnly     = 0
Const msiOpenDatabaseModeTransact     = 1

Const msiViewModifyUpdate  = 2
Const msiReadStreamAnsi    = 2

Dim argCount:argCount = Wscript.Arguments.Count
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If (argCount < 4) Then
	Wscript.Echo "Windows Installer utility to copy a file into a database text field." &_
		vbNewLine & "The 1st argument is the path to the installation database" &_
		vbNewLine & "The 2nd argument is the database table name" &_
		vbNewLine & "The 3rd argument is the set of primary key values, concatenated with colons" &_
		vbNewLine & "The 4th argument is non-key column name to receive the text data" &_
		vbNewLine & "The 5th argument is the path to the text file to copy" &_
		vbNewLine & "If the 5th argument is omitted, the existing data will be listed" &_
		vbNewLine & "All primary keys values must be specified in order, separated by colons" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Connect to Windows Installer object
On Error Resume Next
Dim installer : Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError


' Process input arguments and open database
Dim databasePath: databasePath = Wscript.Arguments(0)
Dim tableName   : tableName    = Wscript.Arguments(1)
Dim rowKeyValues: rowKeyValues = Split(Wscript.Arguments(2),":",-1,vbTextCompare)
Dim dataColumn  : dataColumn   = Wscript.Arguments(3)
Dim openMode : If argCount >= 5 Then openMode = msiOpenDatabaseModeTransact Else openMode = msiOpenDatabaseModeReadOnly
Dim database : Set database = installer.OpenDatabase(databasePath, openMode) : CheckError
Dim keyRecord : Set keyRecord = database.PrimaryKeys(tableName) : CheckError
Dim keyCount : keyCount = keyRecord.FieldCount
If UBound(rowKeyValues) + 1 <> keyCount Then Fail "Incorrect number of primary key values"

' Generate and execute query
Dim predicate, keyIndex
For keyIndex = 1 To keyCount
	If Not IsEmpty(predicate) Then predicate = predicate & " AND "
	predicate = predicate & "`" & keyRecord.StringData(keyIndex) & "`='" & rowKeyValues(keyIndex-1) & "'"
Next
Dim query : query = "SELECT `" & dataColumn & "` FROM `" & tableName & "` WHERE " & predicate
REM Wscript.Echo query 
Dim view : Set view = database.OpenView(query) : CheckError
view.Execute : CheckError
Dim resultRecord : Set resultRecord = view.Fetch : CheckError
If resultRecord Is Nothing Then Fail "Requested table row not present"

' Update value if supplied. Cannot store stream object in string column, must convert stream to string
If openMode = msiOpenDatabaseModeTransact Then
	resultRecord.SetStream 1, Wscript.Arguments(4) : CheckError
	Dim sizeStream : sizeStream = resultRecord.DataSize(1)
	resultRecord.StringData(1) = resultRecord.ReadStream(1, sizeStream, msiReadStreamAnsi) : CheckError
	view.Modify msiViewModifyUpdate, resultRecord : CheckError
	database.Commit : CheckError
Else
	Wscript.Echo resultRecord.StringData(1)
End If

Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
	End If
	Fail message
End Sub

Sub Fail(message)
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIIl7wYJKoZIhvcNAQcCoIIl4DCCJdwCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' +n92KCymBpb0NXQZvmHi9Nqe+zPewtCgF8ON8kB0lL+g
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
'' SIG '' xt8jmpZ1xTGCGcowghnGAgEBMIGVMH4xCzAJBgNVBAYT
'' SIG '' AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
'' SIG '' EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
'' SIG '' cG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2Rl
'' SIG '' IFNpZ25pbmcgUENBIDIwMTACEzMAAARvL2lOUrpvTWQA
'' SIG '' AAAABG8wDQYJYIZIAWUDBAIBBQCgggEEMBkGCSqGSIb3
'' SIG '' DQEJAzEMBgorBgEEAYI3AgEEMBwGCisGAQQBgjcCAQsx
'' SIG '' DjAMBgorBgEEAYI3AgEVMC8GCSqGSIb3DQEJBDEiBCBW
'' SIG '' gYfj52YGyFOL5q3kaUi6lFoP6NuiIn9oa5lut4aLvjA8
'' SIG '' BgorBgEEAYI3CgMcMS4MLHEzZkpUZU1qbTYvcWV0R2JJ
'' SIG '' eVRPR0k0cmwwc1UydFZRUEdTZmx6c1NIcUU9MFoGCisG
'' SIG '' AQQBgjcCAQwxTDBKoCSAIgBNAGkAYwByAG8AcwBvAGYA
'' SIG '' dAAgAFcAaQBuAGQAbwB3AHOhIoAgaHR0cDovL3d3dy5t
'' SIG '' aWNyb3NvZnQuY29tL3dpbmRvd3MwDQYJKoZIhvcNAQEB
'' SIG '' BQAEggEAIHgEB/HJwmLOObjZj5rd+hB+rt6/s+a4OzYf
'' SIG '' tnzfv6Q9Hk/wlfI8dkhzNKKx5FuOqI26D46q64wU8bDJ
'' SIG '' VTVrJtcBt6a0ZCv2vKhTz0NRJydN9nEFMufPdaE+Nc4y
'' SIG '' 41J7bU0y94eJAROvEJWxUJ4pXApVGsEJ0xlNfOp8zMb0
'' SIG '' +6vwH7Xq0haZReAsX/JfvkbB2UDAz6Aj7ni/K65uYnlw
'' SIG '' 5e2pHTWctfu/sTvTTqGG6+UFs4ITSs40ifeqiWhacwX6
'' SIG '' 8hdhfBTnnFXx/FZScIQO882GkPnlO5Hr/NP8hha1KWJJ
'' SIG '' nYlQL3b4Ydd90tLcJAC0NyD8zrk9mlbU/pNRJBoPWKGC
'' SIG '' Fv0wghb5BgorBgEEAYI3AwMBMYIW6TCCFuUGCSqGSIb3
'' SIG '' DQEHAqCCFtYwghbSAgEDMQ8wDQYJYIZIAWUDBAIBBQAw
'' SIG '' ggFRBgsqhkiG9w0BCRABBKCCAUAEggE8MIIBOAIBAQYK
'' SIG '' KwYBBAGEWQoDATAxMA0GCWCGSAFlAwQCAQUABCCQGY9w
'' SIG '' H9NFO1lTcoAlRs6z7vv/4JN3DJGDzNSqQVV5TAIGYrS8
'' SIG '' ELZpGBMyMDIyMDcxNjA4NTY1OC42NjNaMASAAgH0oIHQ
'' SIG '' pIHNMIHKMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
'' SIG '' aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
'' SIG '' ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQL
'' SIG '' ExxNaWNyb3NvZnQgQW1lcmljYSBPcGVyYXRpb25zMSYw
'' SIG '' JAYDVQQLEx1UaGFsZXMgVFNTIEVTTjo0OUJDLUUzN0Et
'' SIG '' MjMzQzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3Rh
'' SIG '' bXAgU2VydmljZaCCEVQwggcMMIIE9KADAgECAhMzAAAB
'' SIG '' lwPPWZxriXg/AAEAAAGXMA0GCSqGSIb3DQEBCwUAMHwx
'' SIG '' CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
'' SIG '' MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
'' SIG '' b3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
'' SIG '' c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTIxMTIw
'' SIG '' MjE5MDUxNFoXDTIzMDIyODE5MDUxNFowgcoxCzAJBgNV
'' SIG '' BAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
'' SIG '' VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
'' SIG '' Q29ycG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBB
'' SIG '' bWVyaWNhIE9wZXJhdGlvbnMxJjAkBgNVBAsTHVRoYWxl
'' SIG '' cyBUU1MgRVNOOjQ5QkMtRTM3QS0yMzNDMSUwIwYDVQQD
'' SIG '' ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIIC
'' SIG '' IjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKCAgEA7QBK
'' SIG '' 6kpBTfPwnv3LKx1VnL9YkozUwKzyhDKij1E6WCV/EwWZ
'' SIG '' fPCza6cOGxKT4pjvhLXJYuUQaGRInqPks2FJ29PpyhFm
'' SIG '' hGILm4Kfh0xWYg/OS5Xe5pNl4PdSjAxNsjHjiB9gx6U7
'' SIG '' J+adC39Ag5XzxORzsKT+f77FMTXg1jFus7ErilOvWi+z
'' SIG '' nMpN+lTMgioxzTC+u1ZmTCQTu219b2FUoTr0KmVJMQqQ
'' SIG '' kd7M5sR09PbOp4cC3jQs+5zJ1OzxIjRlcUmLvldBE6aR
'' SIG '' aSu0x3BmADGt0mGY0MRsgznOydtJBLnerc+QK0kcxuO6
'' SIG '' rHA3z2Kr9fmpHsfNcN/eRPtZHOLrpH59AnirQA7puz6k
'' SIG '' a20TA+8MhZ19hb8msrRo9LmirjFxSbGfsH3ZNEbLj3lh
'' SIG '' 7Vc+DEQhMH2K9XPiU5Jkt5/6bx6/2/Od3aNvC6Dx3s5N
'' SIG '' 3UsW54kKI1twU2CS5q1Hov5+ARyuZk0/DbsRus6D97fB
'' SIG '' 1ZoQlv/4trBcMVRz7MkOrHa8bP4WqbD0ebLYtiExvx4H
'' SIG '' uEnh+0p3veNjh3gP0+7DkiVwIYcfVclIhFFGsfnSiFex
'' SIG '' ruu646uUla+VTUuG3bjqS7FhI3hh6THov/98XfHcWeNh
'' SIG '' vxA5K+fi+1BcSLgQKvq/HYj/w/Mkf3bu73OERisNaaca
'' SIG '' aOCR/TJ2H3fs1A7lIHECAwEAAaOCATYwggEyMB0GA1Ud
'' SIG '' DgQWBBRtzwHPKOswbpZVC9Gxvt1+vRUAYDAfBgNVHSME
'' SIG '' GDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8E
'' SIG '' WDBWMFSgUqBQhk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
'' SIG '' b20vcGtpb3BzL2NybC9NaWNyb3NvZnQlMjBUaW1lLVN0
'' SIG '' YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUH
'' SIG '' AQEEYDBeMFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1p
'' SIG '' Y3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY3Jvc29m
'' SIG '' dCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNy
'' SIG '' dDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsGAQUF
'' SIG '' BwMIMA0GCSqGSIb3DQEBCwUAA4ICAQAESNhh0iTtMx57
'' SIG '' IXLfh4LuHbD1NG9MlLA1wYQHQBnR9U/rg3qt3Nx6e7+Q
'' SIG '' uEKMEhKqdLf3g5RR4R/oZL5vEJVWUfISH/oSWdzqrShq
'' SIG '' cmT4Oxzc2CBs0UtnyopVDm4W2Cumo3quykYPpBoGdeir
'' SIG '' vDdd153AwsJkIMgm/8sxJKbIBeT82tnrUngNmNo8u7l1
'' SIG '' uE0hsMAq1bivQ63fQInr+VqYJvYT0W/0PW7pA3qh4ocN
'' SIG '' jiX6Z8d9kjx8L7uBPI/HsxifCj/8mFRvpVBYOyqP7Y5d
'' SIG '' i5ZAnjTDSHMZNUFPHt+nhFXUcHjXPRRHCMqqJg4D63X6
'' SIG '' b0V0R87Q93ipwGIXBMzOMQNItJORekHtHlLi3bg6Lnpj
'' SIG '' s0aCo5/RlHCjNkSDg+xV7qYea37L/OKTNjqmH3pNAa3B
'' SIG '' vP/rDQiGEYvgAbVHEIQz7WMWSYsWeUPFZI36mCjgUY6V
'' SIG '' 538CkQtDwM8BDiAcy+quO8epykiP0H32yqwDh852BeWm
'' SIG '' 1etF+Pkw/t8XO3Q+diFu7Ggiqjdemj4VfpRsm2tTN9Hn
'' SIG '' Aewrrb0XwY8QE2tp0hRdN2b0UiSxMmB4hNyKKXVaDLOF
'' SIG '' CdiLnsfpD0rjOH8jbECZObaWWLn9eEvDr+QNQPvS4r47
'' SIG '' L9Aa8Lr1Hr47VwJ5E2gCEnvYwIRDzpJhMRi0KijYN43y
'' SIG '' T6XSGR4N9jCCB3EwggVZoAMCAQICEzMAAAAVxedrngKb
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
'' SIG '' ZvKhggLLMIICNAIBATCB+KGB0KSBzTCByjELMAkGA1UE
'' SIG '' BhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
'' SIG '' BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
'' SIG '' b3Jwb3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFt
'' SIG '' ZXJpY2EgT3BlcmF0aW9uczEmMCQGA1UECxMdVGhhbGVz
'' SIG '' IFRTUyBFU046NDlCQy1FMzdBLTIzM0MxJTAjBgNVBAMT
'' SIG '' HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoB
'' SIG '' ATAHBgUrDgMCGgMVAGFA0rCNmEk0zU12DYNGMU3B1mPR
'' SIG '' oIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
'' SIG '' Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAc
'' SIG '' BgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
'' SIG '' A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIw
'' SIG '' MTAwDQYJKoZIhvcNAQEFBQACBQDmfOMkMCIYDzIwMjIw
'' SIG '' NzE2MTUxMTMyWhgPMjAyMjA3MTcxNTExMzJaMHQwOgYK
'' SIG '' KwYBBAGEWQoEATEsMCowCgIFAOZ84yQCAQAwBwIBAAIC
'' SIG '' HCUwBwIBAAICEecwCgIFAOZ+NKQCAQAwNgYKKwYBBAGE
'' SIG '' WQoEAjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAweh
'' SIG '' IKEKMAgCAQACAwGGoDANBgkqhkiG9w0BAQUFAAOBgQAR
'' SIG '' MdT0eZbsDKHbqmilc3L5Sa8rgnjbpOjONblfLXOk2TX3
'' SIG '' B3OBRXtv72jP/74uqwMDAZZImr/laH75A9t97MC0qBbL
'' SIG '' 8uTubp9BZedwJIXUpZz9zzJc2ZAYrgYmMbwQq/oRjZdx
'' SIG '' 2ASptYHWVVH6f5iE9NyQ0fgnJoz0E7euSIIC8TGCBA0w
'' SIG '' ggQJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQI
'' SIG '' EwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4w
'' SIG '' HAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAk
'' SIG '' BgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAy
'' SIG '' MDEwAhMzAAABlwPPWZxriXg/AAEAAAGXMA0GCWCGSAFl
'' SIG '' AwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMxDQYLKoZIhvcN
'' SIG '' AQkQAQQwLwYJKoZIhvcNAQkEMSIEILO4OR0u9umnsY5U
'' SIG '' t0/1jdSX0zqIgYthEz80R+ZTV32DMIH6BgsqhkiG9w0B
'' SIG '' CRACLzGB6jCB5zCB5DCBvQQgW3vaGxCVejj+BAzFSfMx
'' SIG '' fHQ+bxxkqCw8LkMY/QZ4pr8wgZgwgYCkfjB8MQswCQYD
'' SIG '' VQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
'' SIG '' A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
'' SIG '' IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
'' SIG '' VGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAAZcDz1mca4l4
'' SIG '' PwABAAABlzAiBCA57obvITwCiIoAFiidTdXAMsvKi/wK
'' SIG '' zVgKwzNngtPq4DANBgkqhkiG9w0BAQsFAASCAgCsaRm+
'' SIG '' aD/iZL2zLSK26tpkPpSURDWP8gIem8Gxeax6YnVdBH9b
'' SIG '' 0f3TitGgfDL6KBPXhYXafgFY9Y+gNrERCn1MNwPkWjDF
'' SIG '' p8dFgHA0GaxpyXp8c9uTirVMgOypwpkw9iAHtgQd6iyE
'' SIG '' /aotKy/Z31No8jAomogNKXzSkRafCwX3TpxTnRh6+Y4/
'' SIG '' HHBhh7IBh29YB2RhTxF1ThAd0CaZce4W7X/GMJ3c4r7C
'' SIG '' WxNSysRiy9NRN51NlVi17dOqhYZeNNC3HQmGXTEle/N6
'' SIG '' nKaTaPYQ2UzXyg7AlTxU6dGpGlopOj+Ls20+ccVY6FzX
'' SIG '' 8SNWopeetrJp2T9GSBH43JTOvyvSoRQ0GIIZG9s5JRJ1
'' SIG '' EGuIsP6cOQGJSsnM+pnDm+4+PZcF0i7kqLX6Lt0H6kMQ
'' SIG '' X8j0S2oZRqMaVOha9A9SUSDKVvTczFrTtBuedBY1A4Jd
'' SIG '' 8dM4R4twIDbElCm2GQhig9FucORoR9BSc2C443fF1vIW
'' SIG '' pIVPIgWQz0eiLy1vsjOgOVS7z4VSWyb6yZHLuYrSB9SJ
'' SIG '' ZOzS/eMeRlzGVjhjeDvxprtgytUY9oiB5PO82k0TqjDG
'' SIG '' JM3YLRvvEVuDJAVTGTLrfPPHUQKFIv6nfEykLXwnxCrM
'' SIG '' K6f3OCiHXPzrNL0Dxtm70d2FozITbA0rP+y+hiNzXK5i
'' SIG '' 5maVLb+UX31XZVJ66w==
'' SIG '' End signature block

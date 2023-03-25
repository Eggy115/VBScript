' Windows Installer utility to list component composition of an MSI database
' For use with Windows Scripting Host, CScript.exe or WScript.exe
' Copyright (c) Microsoft Corporation. All rights reserved.
' Demonstrates the various tables having foreign keys to the Component table
'
Option Explicit
Public isGUI, installer, database, message, compParam  'global variables access across functions

Const msiOpenDatabaseModeReadOnly     = 0

' Check if run from GUI script host, in order to modify display
If UCase(Mid(Wscript.FullName, Len(Wscript.Path) + 2, 1)) = "W" Then isGUI = True

' Show help if no arguments or if argument contains ?
Dim argCount:argCount = Wscript.Arguments.Count
If argCount > 0 Then If InStr(1, Wscript.Arguments(0), "?", vbTextCompare) > 0 Then argCount = 0
If argCount = 0 Then
	Wscript.Echo "Windows Installer utility to list component composition in an install database." &_
		vbLf & " The 1st argument is the path to an install database, relative or complete path" &_
		vbLf & " The 2nd argument is the name of the component (primary key of Component table)" &_
		vbLf & " If the 2nd argument is not present, the names of all components will be listed" &_
		vbLf & " If the 2nd argument is a ""*"", the composition of all components will be listed" &_
		vbLf & " Large databases or components are better displayed using CScript than WScript." &_
		vbLf & " Note: The name of the component, if provided,  is case-sensitive" &_
		vbNewLine &_
		vbNewLine & "Copyright (C) Microsoft Corporation.  All rights reserved."
	Wscript.Quit 1
End If

' Connect to Windows Installer object
On Error Resume Next
Set installer = Nothing
Set installer = Wscript.CreateObject("WindowsInstaller.Installer") : CheckError

' Open database
Dim databasePath:databasePath = Wscript.Arguments(0)
Set database = installer.OpenDatabase(databasePath, msiOpenDatabaseModeReadOnly) : CheckError

If argCount = 1 Then  'If no component specified, then simply list components
	ListComponents False
	ShowOutput "Components for " & databasePath, message
ElseIf Left(Wscript.Arguments(1), 1) = "*" Then 'List all components
	ListComponents True
Else
	QueryComponent Wscript.Arguments(1) 
End If
Wscript.Quit 0

' List all table rows referencing a given component
Function QueryComponent(component)
	' Get component info and format output header
	Dim view, record, header, componentId
	Set view = database.OpenView("SELECT `ComponentId` FROM `Component` WHERE `Component` = ?") : CheckError
	Set compParam = installer.CreateRecord(1)
	compParam.StringData(1) = component
	view.Execute compParam : CheckError
	Set record = view.Fetch : CheckError
	Set view = Nothing
	If record Is Nothing Then Fail "Component not in database: " & component
	componentId = record.StringData(1)
	header = "Component: "& component & "  ComponentId = " & componentId

	' List of tables with foreign keys to Component table - with subsets of columns to display
	DoQuery "FeatureComponents","Feature_"                           '
	DoQuery "PublishComponent", "ComponentId,Qualifier"              'AppData,Feature
	DoQuery "File",             "File,Sequence,FileName,Version"     'FileSize,Language,Attributes
	DoQuery "SelfReg,File",     "File_"                              'Cost
	DoQuery "BindImage,File",   "File_"                              'Path
	DoQuery "Font,File",        "File_,FontTitle"                    '
	DoQuery "Patch,File",       "File_"                              'Sequence,PatchSize,Attributes,Header
	DoQuery "DuplicateFile",    "FileKey,File_,DestName"             'DestFolder
	DoQuery "MoveFile",         "FileKey,SourceName,DestName"        'SourceFolder,DestFolder,Options
	DoQuery "RemoveFile",       "FileKey,FileName,DirProperty"       'InstallMode
	DoQuery "IniFile",          "IniFile,FileName,Section,Key"       'Value,Action
	DoQuery "RemoveIniFile",    "RemoveIniFile,FileName,Section,Key" 'Value,Action
	DoQuery "Registry",         "Registry,Root,Key,Name"             'Value
	DoQuery "RemoveRegistry",   "RemoveRegistry,Root,Key,Name"       '
	DoQuery "Shortcut",         "Shortcut,Directory_,Name,Target"    'Arguments,Description,Hotkey,Icon_,IconIndex,ShowCmd,WkDir
	DoQuery "Class",            "CLSID,Description"                  'Context,ProgId_Default,AppId_,FileType,Mask,Icon_,IconIndex,DefInprocHandler,Argument,Feature_
	DoQuery "ProgId,Class",     "Class_,ProgId,Description"          'ProgId_Parent,Icon_IconIndex,Insertable
	DoQuery "Extension",        "Extension,ProgId_"                  'MIME_,Feature_
	DoQuery "Verb,Extension",   "Extension_,Verb"                    'Sequence,Command.Argument
	DoQuery "MIME,Extension",   "Extension_,ContentType"             'CLSID
	DoQuery "TypeLib",          "LibID,Language,Version,Description" 'Directory_,Feature_,Cost
	DoQuery "CreateFolder",     "Directory_"                         ' 
	DoQuery "Environment",      "Environment,Name"                   'Value
	DoQuery "ODBCDriver",       "Driver,Description"                 'File_,File_Setup
	DoQuery "ODBCAttribute,ODBCDriver", "Driver_,Attribute,Value" '
	DoQuery "ODBCTranslator",   "Translator,Description"             'File_,File_Setup
	DoQuery "ODBCDataSource",   "DataSource,Description,DriverDescription" 'Registration
	DoQuery "ODBCSourceAttribute,ODBCDataSource", "DataSource_,Attribute,Value" '
	DoQuery "ServiceControl",   "ServiceControl,Name,Event"          'Arguments,Wait
	DoQuery "ServiceInstall",   "ServiceInstall,Name,DisplayName"    'ServiceType,StartType,ErrorControl,LoadOrderGroup,Dependencies,StartName,Password
	DoQuery "ReserveCost",      "ReserveKey,ReserveFolder"           'ReserveLocal,ReserveSource

	QueryComponent = ShowOutput(header, message)
	message = Empty
End Function

' List all components in database
Sub ListComponents(queryAll)
	Dim view, record, component
	Set view = database.OpenView("SELECT `Component`,`ComponentId` FROM `Component`") : CheckError
	view.Execute : CheckError
	Do
		Set record = view.Fetch : CheckError
		If record Is Nothing Then Exit Do
		component = record.StringData(1)
		If queryAll Then
			If QueryComponent(component) = vbCancel Then Exit Sub
		Else
			If Not IsEmpty(message) Then message = message & vbLf
			message = message & component
		End If
	Loop
End Sub

' Perform a join to query table rows linked to a given component, delimiting and qualifying names to prevent conflicts
Sub DoQuery(table, columns)
	Dim view, record, columnCount, column, output, header, delim, columnList, tableList, tableDelim, query, joinTable, primaryKey, foreignKey, columnDelim
	On Error Resume Next
	tableList  = Replace(table,   ",", "`,`")
	tableDelim = InStr(1, table, ",", vbTextCompare)
	If tableDelim Then  ' need a 3-table join
		joinTable = Right(table, Len(table)-tableDelim)
		table = Left(table, tableDelim-1)
		foreignKey = columns
		Set record = database.PrimaryKeys(joinTable)
		primaryKey = record.StringData(1)
		columnDelim = InStr(1, columns, ",", vbTextCompare)
		If columnDelim Then foreignKey = Left(columns, columnDelim - 1)
		query = " AND `" & foreignKey & "` = `" & primaryKey & "`"
	End If
	columnList = table & "`." & Replace(columns, ",", "`,`" & table & "`.`")
	query = "SELECT `" & columnList & "` FROM `" & tableList & "` WHERE `Component_` = ?" & query
	If database.TablePersistent(table) <> 1 Then Exit Sub
	Set view = database.OpenView(query) : CheckError
	view.Execute compParam : CheckError
	Do
		Set record = view.Fetch : CheckError
		If record Is Nothing Then Exit Do
		If IsEmpty(output) Then
			If Not IsEmpty(message) Then message = message & vbLf
			message = message & "----" & table & " Table----  (" & columns & ")" & vbLf
		End If
		output = Empty
		columnCount = record.FieldCount
		delim = "  "
		For column = 1 To columnCount
			If column = columnCount Then delim = vbLf
			output = output & record.StringData(column) & delim
		Next
		message = message & output
	Loop
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

Function ShowOutput(header, message)
	ShowOutput = vbOK
	If IsEmpty(message) Then Exit Function
	If isGUI Then
		ShowOutput = MsgBox(message, vbOKCancel, header)
	Else
		Wscript.Echo "> " & header
		Wscript.Echo message
	End If
End Function

Sub Fail(message)
	Wscript.Echo message
	Wscript.Quit 2
End Sub

'' SIG '' Begin signature block
'' SIG '' MIIl9gYJKoZIhvcNAQcCoIIl5zCCJeMCAQExDzANBglg
'' SIG '' hkgBZQMEAgEFADB3BgorBgEEAYI3AgEEoGkwZzAyBgor
'' SIG '' BgEEAYI3AgEeMCQCAQEEEE7wKRaZJ7VNj+Ws4Q8X66sC
'' SIG '' AQACAQACAQACAQACAQAwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' WoLLQA6rHA8fttRtGpZpVGF985uNg+TIhlmzKb0W2sOg
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
'' SIG '' IgQgU+H2zhtPO/bs1TiBZyzIPq79ccTVGaHmw1wCd9Wp
'' SIG '' jFgwPAYKKwYBBAGCNwoDHDEuDCxjRG5BMkVsM0ZYSFls
'' SIG '' Y0JLVW5FMzFFOEtmZVpRNEZNSldoNGtvR2g1YVlzPTBa
'' SIG '' BgorBgEEAYI3AgEMMUwwSqAkgCIATQBpAGMAcgBvAHMA
'' SIG '' bwBmAHQAIABXAGkAbgBkAG8AdwBzoSKAIGh0dHA6Ly93
'' SIG '' d3cubWljcm9zb2Z0LmNvbS93aW5kb3dzMA0GCSqGSIb3
'' SIG '' DQEBAQUABIIBAFsKbcpDJJqsU9ghPGU58DKR3n0soh3k
'' SIG '' TScN/lgOgNUIbc476bQ7qzUyffR10Ysdq0PiiOD1d+H4
'' SIG '' eJa/BqB1ua03GUUqSOQbPGqS1QUE+3cYTpW/n3I4Z2RV
'' SIG '' ZvXSJnmr00VdM4AjT2LHbkAbinfm/09KNdDv3HV11RGq
'' SIG '' DdjM+ddHtqIzX+TkCXIlhxyUKQH3yWqh0V0obL+L50KF
'' SIG '' PzKEkXy/3mZoj/OjLCxV67TcsrncsBKfQ+3S3TTr33dx
'' SIG '' 6mEz3I0KQG9tPnmXJH8gccO+bv6KeBqoZnV8RRCzt8Me
'' SIG '' 9RGqJNd6ScKdpwVpKun/bmJnSD9mEmoErCckwJAjL5G9
'' SIG '' S9qhghcAMIIW/AYKKwYBBAGCNwMDATGCFuwwghboBgkq
'' SIG '' hkiG9w0BBwKgghbZMIIW1QIBAzEPMA0GCWCGSAFlAwQC
'' SIG '' AQUAMIIBUQYLKoZIhvcNAQkQAQSgggFABIIBPDCCATgC
'' SIG '' AQEGCisGAQQBhFkKAwEwMTANBglghkgBZQMEAgEFAAQg
'' SIG '' Za9LJDPNZPNHbILZuAtftU5tUBdU+cUXVU65qyNeAJgC
'' SIG '' BmLP9QXTPBgTMjAyMjA3MTYwODU2NTguNjM2WjAEgAIB
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
'' SIG '' CyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCCy5hO/
'' SIG '' 45PzTUBoCbXA6muPSjxrEDOHaP/Ogw+I1is6czCB+gYL
'' SIG '' KoZIhvcNAQkQAi8xgeowgecwgeQwgb0EIDcPRYUgjSzK
'' SIG '' OhF39d4QgbRZQgrPO7Lo/qE5GtvSeqa8MIGYMIGApH4w
'' SIG '' fDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
'' SIG '' b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1p
'' SIG '' Y3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
'' SIG '' cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAGc
'' SIG '' D6ZNYdKeSygAAQAAAZwwIgQggEMpQEe+SDkVnD6hE7l2
'' SIG '' h1gRNE1rq4Oh79tMwbg9E8UwDQYJKoZIhvcNAQELBQAE
'' SIG '' ggIAsHciVj7HYJzkoglbtR6WmfXkjq6Ar/IauxdTK8Cd
'' SIG '' Rn8FZuSoa+iJ/MYntwojc3qgwB6XPiRBDUIId4LMa2aI
'' SIG '' TqKz6FYBeLMvMnqVwKse9YX3hbFFkPx7HQ/haIscvICo
'' SIG '' 9Jm6RKSeQ90YeHTFKA/R0s7262KuLo2bWdz3SPhKfgX8
'' SIG '' ktVuNcsFRzHHSdchoP27Y53TQG4Y5Q+kc2qzpJAjp4A3
'' SIG '' zDHj4cV7h6jvPJ/16IH0aGT9Fh9QswSKwpW37uo9kAWG
'' SIG '' ijNBR0yxy8x3/ur9OvWmwkmXpcLBF88nnxHRzWPwGXNs
'' SIG '' HuakIU9cyj/PjvBuU9TV/V0h+8EQ9SR7Z2pvaRnhiQ4Z
'' SIG '' 62mv2kQlcQgXVQl3a7WBm5rnYoxP+NJU0Wpg3BwrOE1d
'' SIG '' KYgteWar9tj+JnvoRfcGbP1jj8Ym5V3IumRB4WBbop/e
'' SIG '' KOewm3u8JtJi4x/t9Jzhh3Yhcu/mqai8NVUf5usX2Agi
'' SIG '' 2NbFhIrVZN+Q1zvZJnka4CxEVYU/oH+qGVST+tFXWxqe
'' SIG '' SnMpBwVhs1DNy7sJIthwafGY42MVt0IbnWPgkMeZFnBJ
'' SIG '' woGOs4IdXLxGVqQh/pWiXLsuC2ai5gt0NipPl++K4BGB
'' SIG '' rK9QvwEbs5j++8QiXt0OzGb1XVzM8fikCbbU06mxLEkJ
'' SIG '' 8FYdpQlGh+iv4189S+7nv4t+xIs=
'' SIG '' End signature block

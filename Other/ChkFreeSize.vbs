Option Explicit

Dim FSO, File, FileFreeSizeChk, FileWarning, FileDrvSpace, FileRPSize, FSOFreeSizeChk
Dim SizeFreeSpace, SizeMSCSpace, SizeRP, SizePageFile, SizeHiberFile, SizeRest, SizeWarning
Dim line, Result, FFS
Dim OSDskIdx, OSDskSize
Dim objWMI, colItems, objItem
Dim FoundSection
Dim bResult
Const ForReading = 1
Const ForAppending = 8

Set FSO = CreateObject("Scripting.FileSystemObject")
SizeFreeSpace = 0
SizeMSCSpace = 0
SizeRP = 0
SizePageFile = 0
SizeHiberFile = 0
SizeRest = 0
SizeWarning = 0
FFS = False
OSDskIdx = 0
OSDskSize = 0
FoundSection = 0
FileFreeSizeChk = "C:\System.sav\Logs\ChkFreeSize.log"
FileWarning = "C:\System.sav\Logs\Warning.log"
FileDrvSpace = "C:\System.sav\Logs\DrvSpace.log"
FileRPSize = "C:\System.sav\Logs\RPSize.ini"

' ===== Open log file to be written =====
Set FSOFreeSizeChk = FSO.CreateTextFile(FileFreeSizeChk, True)

' ===== Get free space & MSC size =====
If FSO.FileExists(FileDrvSpace) = False Then
	FSOFreeSizeChk.WriteLine ""
	FSOFreeSizeChk.WriteLine "Cannot find " & FileDrvSpace
	FSOFreeSizeChk.WriteLine ""
	FSOFreeSizeChk.Close
	Wscript.Quit(1)
End If

Set File = FSO.OpenTextFile(FileDrvSpace, ForReading)
Do Until File.AtEndOfStream
	line = File.ReadLine

	If InStr(1, line, "[Beginning of 2nd Capture stage]") > 0 Then
		FoundSection = 1
	End If

	If FoundSection = 1 Then
		If InStr(1, line,"Free:") > 0 Then
			If SizeFreeSpace <> 0 Then
				Exit Do
			End If

			Result = Split(line, " ")
			SizeFreeSpace = CDbl(Trim(Result(1)))
		End If

		If InStr(1, line,"Size:") > 0 Then
			If SizeMSCSpace <> 0 Then
				Exit Do
			End If

			Result = Split(line, " ")
			SizeMSCSpace = CDbl(Trim(Result(1)))
		End If

		If InStr(1, line,"hiberfil.sys:") > 0 Then
			Result = Split(line, " ")
			SizeHiberFile = CDbl(Trim(Result(1)))
		End If

		If InStr(1, line,"pagefile.sys:") > 0 Then
			Result = Split(line, " ")
			SizePageFile = CDbl(Trim(Result(1)))
		End If
	End If
Loop

' ===== Check if OS disk has FFS partition =====
Set objWMI = GetObject("winmgmts:\\.\root\Microsoft\Windows\Storage")
Set colItems = objWMI.ExecQuery("SELECT * FROM MSFT_Partition")

For Each objItem In colItems
	If LCase(objItem.GptType) = "{d3bfe2de-3daf-11df-ba40-e3a556d89593}" Then
		If objItem.PartitionNumber > 1 Then
			FFS = True
			OSDskIdx = objItem.DiskNumber
			Exit For
		End If
	End If
Next

If FFS = True Then
	' ***************************
	' **   OS disk with FFS    **
	' ***************************

	' ===== Get OS disk size =====
	Set colItems = objWMI.ExecQuery("SELECT * FROM MSFT_Disk WHERE Number=" & OSDskIdx)

	For Each objItem In colItems
		OSDskSize = objItem.Size
		Exit For
	Next

	' ===== Get RP size =====
	If FSO.FileExists(FileRPSize) = False Then
		FSOFreeSizeChk.WriteLine ""
		FSOFreeSizeChk.WriteLine "Cannot find " & FileRPSize
		FSOFreeSizeChk.WriteLine ""
		FSOFreeSizeChk.Close
		Wscript.Quit(1)
	End If

	Set File = FSO.OpenTextFile(FileRPSize, ForReading)
	Do Until File.AtEndOfStream
		line = File.ReadLine
		If InStr(1, line, "RPSize") > 0 Then
			Result = Split(line, "=")
			SizeRP = Trim(Result(1))
			Exit Do
		End If
	Loop

	' ===== Calculate rest free size =====
	SizeRest = (OSDskSize - SizeMSCSpace) / 1024 / 1024 / 1024 - (789 + SizeRP) / 1024

	' ===== Write logs =====
	FSOFreeSizeChk.WriteLine ""
	FSOFreeSizeChk.WriteLine "====================================="
	FSOFreeSizeChk.WriteLine " Free on SSD w/ FFS "
	FSOFreeSizeChk.WriteLine "====================================="
	FSOFreeSizeChk.WriteLine "OS disk size = " & OSDskSize & " B"
	FSOFreeSizeChk.WriteLine "MSC size = " & SizeMSCSpace & " B"
	FSOFreeSizeChk.WriteLine "RP size = " & SizeRP & " MB"
	FSOFreeSizeChk.WriteLine ""
	FSOFreeSizeChk.WriteLine "4G: " & (SizeRest - 4) & " GB"
	FSOFreeSizeChk.WriteLine "8G: " & (SizeRest - 8) & " GB"
	FSOFreeSizeChk.WriteLine ""
	FSOFreeSizeChk.Close

	If SizeRest <= 8 Then
		Call CopyFileContent(FileFreeSizeChk, FileWarning)
	End If
Else
	' ***************************
	' **  OS disk without FFS  **
	' ***************************

	' ===== Get page file size =====
	bResult = 0
	Set File = FSO.OpenTextFile("C:\System.sav\Flags\PINCTRLTwk.pf", ForReading)
	Do Until File.AtEndOfStream
		line = File.ReadLine
		If InStr(1, line, "SetPF") > 0 Then
			Result = Split(line, "=")
			bResult = Trim(Result(1))
			Exit Do
		End If
	Loop

	If bResult = 0 Then
		Set File = FSO.GetFile("C:\pagefile.sys")
		SizePageFile = File.Size
	Else
		SizePageFile = 1000 * 1024 * 1024
	End If

	' ===== Get hibernation file size =====
	If FSO.FileExists("C:\System.sav\Flags\WimBoot.flg") = False Then
		Set File = FSO.GetFile("C:\hiberfil.sys")
		SizeHiberFile = File.Size
	End If

	' ===== Calculate rest free size =====
	SizeRest = (SizeFreeSpace + SizePageFile + SizeHiberFile) / 1024 / 1024 / 1024

	' ===== Write logs =====
	' Pagefile = memory * 100%
	' Hiberfile = memory * 75%
	FSOFreeSizeChk.WriteLine ""
	FSOFreeSizeChk.WriteLine "====================================="
	FSOFreeSizeChk.WriteLine " Free on Defined MSC w/ 4/8/16GB RAM "
	FSOFreeSizeChk.WriteLine "====================================="
	FSOFreeSizeChk.WriteLine "UP free space = " & SizeFreeSpace & " B"
	FSOFreeSizeChk.WriteLine "Page file size = " & SizePageFile & " B"
	FSOFreeSizeChk.WriteLine "Hibernation file size = " & SizeHiberFile & " B"
	FSOFreeSizeChk.WriteLine ""

	If FSO.FileExists("C:\System.sav\Flags\WimBoot.flg") = False Then
		FSOFreeSizeChk.WriteLine "4G: " & (SizeRest - 9) & " GB"
		FSOFreeSizeChk.WriteLine "8G: " & (SizeRest - 16) & " GB"
		FSOFreeSizeChk.WriteLine "16G: " & (SizeRest - 30) & " GB"
		SizeWarning = 30
	Else
		FSOFreeSizeChk.WriteLine "4G: " & (SizeRest - 2) & " GB"
		FSOFreeSizeChk.WriteLine "8G: " & (SizeRest - 2) & " GB"
		FSOFreeSizeChk.WriteLine "16G: " & (SizeRest - 2) & " GB"
		SizeWarning = 2
	End If

	FSOFreeSizeChk.WriteLine ""
	FSOFreeSizeChk.Close

	If SizeRest <= SizeWarning Then
		Call CopyFileContent(FileFreeSizeChk, FileWarning)
	End If
End If

Wscript.Quit(1)

Sub CopyFileContent(FileSrc, FileDest)
	Dim FSOSrc, FSODest

	If FSO.FileExists(FileSrc) = False Then
		Exit Sub
	End If

	Set FSOSrc = FSO.OpenTextFile(FileSrc, ForReading)
	Set FSODest = FSO.OpenTextFile(FileDest, ForAppending, True)

	FSODest.Write FSOSrc.ReadAll
	FSOSrc.Close
	FSODest.Close
End Sub
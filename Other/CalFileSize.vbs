' ===================================================
' Query File Size by Byte
' Date : 2014-04-22
' Usage: CalcFileSize.vbs FileFullPath
' ===================================================
' const 
const ForReading = 1, ForWriting = 2, ForAppending = 3
const MB=1048576

Set objArgs = Wscript.Arguments
If objArgs.Count <> 1 then
   WScript.ECHO "CalFileSize.vbs error args"
   WSCript.Quit 1
End If

strFilePath = Wscript.Arguments(0)
OneFileSize = Round(QueryFileSizeByte(strFilePath))

WScript.Echo  "FileSize.BYTE=" & OneFileSize
WScript.Echo  "FileSize.KB="   & Round(OneFileSize / 1024 )
WScript.Echo  "FileSize.MB="   & Round(OneFileSize / 1024 / 1024 )
' ===================================================
' Query A File Size
' ===================================================
Function QueryFileSizeByte(strFilePath)
	strFilePath = Trim(Replace( strFilePath, "	", ""))
	QueryFileSizeByte = 0
	If CheckIfFileExist(strFilePath) Then
		Set fs = CreateObject("Scripting.FileSystemObject")
		set OneFile = fs.GetFile(strFilePath)
		QueryFileSizeByte = OneFile.Size
	End If
End Function



' ===================================================
' Check if the file was exist
' ===================================================
Function CheckIfFileExist(strFilePath)
	strFilePath = Trim(Replace( strFilePath, "	", ""))
	CheckIfFileExist = 1
	Set fs = CreateObject("Scripting.FileSystemObject")
	If not fs.fileExists(strFilePath) Then
		CheckIfFileExist = 0
	End If
End Function



' List Fax Server Incoming Archive Information


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objFolder = objFaxServer.Folders

Set objIncomingArchive = objFolder.IncomingArchive
Wscript.Echo "Age limit: " & objIncomingArchive.AgeLimit
Wscript.Echo "Archive folder: " & objIncomingArchive.ArchiveFolder
Wscript.Echo "High quota watermark: " & objIncomingArchive.HighQuotaWatermark
Wscript.Echo "Low quota watermark: " & objIncomingArchive.LowQuotaWatermark
Wscript.Echo "Size low: " & objIncomingArchive.SizeLow
Wscript.Echo "Size high: " & objIncomingArchive.SizeHigh
Wscript.Echo "Size quota warning: " & objIncomingArchive.SizeQuotaWarning
Wscript.Echo "Use archive: " & objIncomingArchive.UseArchive

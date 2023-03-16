' List Fax Server Outgoing Archive Information


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objFolder = objFaxServer.Folders
Set objOutgoingArchive = objFolder.OutgoingArchive

Wscript.Echo "Age limikt: " & objOutgoingArchive.AgeLimit
Wscript.Echo "Archive folder: " & objOutgoingArchive.ArchiveFolder
Wscript.Echo "High quota watermark: " & objOutgoingArchive.HighQuotaWatermark
Wscript.Echo "Low quota watermark: " & objOutgoingArchive.LowQuotaWatermark
Wscript.Echo "Size low: " & objOutgoingArchive.SizeLow
Wscript.Echo "Size high: " & objOutgoingArchive.SizeHigh
Wscript.Echo "Size quota warning: " & objOutgoingArchive.SizeQuotaWarning
Wscript.Echo "Use archive: " & objOutgoingArchive.UseArchive

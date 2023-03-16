' List Fax Server Outgoing Queue Information


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objFolder = objFaxServer.Folders
Set objOutgoingQueue = objFolder.OutgoingQueue

Wscript.Echo "Age limit: " & objOutgoingQueue.AgeLimit
Wscript.Echo "Allow personal cover pages: " & _
    objOutgoingQueue.AllowPersonalCoverPages
Wscript.Echo "Blocked: " & objOutgoingQueue.Blocked
Wscript.Echo "Branding: " & objOutgoingQueue.Branding
Wscript.Echo "Discount rate end: " & objOutgoingQueue.DiscountRateEnd
Wscript.Echo "Discount rate start: " & objOutgoingQueue.DiscountRateStart
Wscript.Echo "Paused: " & objOutgoingQueue.Paused
Wscript.Echo "Retries: " & objOutgoingQueue.Retries
Wscript.Echo "Retry delay: " & objOutgoingQueue.RetryDelay
Wscript.Echo "Use Device TSID: " & objOutgoingQueue.UseDeviceTSID

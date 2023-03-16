


' List Fax Server Incoming Queue Information


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objFolder = objFaxServer.Folders

Set objIncomingQueue = objFolder.IncomingQueue
Wscript.Echo "Blocked: " & objIncomingQueue.Blocked

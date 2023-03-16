' List Fax Server Device Information


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set colDevices = objFaxServer.GetDevices()

For Each objFaxDevice in colDevices
    Wscript.Echo "ID: " & objFaxDevice.ID
    Wscript.Echo "CSID: " & objFaxDevice.CSID
    Wscript.Echo "Description: " & objFaxDevice.Description
    Wscript.Echo "Device name: " & objFaxDevice.DeviceName
    Wscript.Echo "Powered off: " & objFaxDevice.PoweredOff
    Wscript.Echo "Provider unique name: " & _
        objFaxDevice.ProviderUniqueName
    Wscript.Echo "Receive mode: " & objFaxDevice.ReceiveMode
    Wscript.Echo "Receiving now: " & objFaxDevice.ReceivingNow
    Wscript.Echo "Ringing now: " & objFaxDevice.RingingNow
    Wscript.Echo "Rings before answer: " & _
        objFaxDevice.RingsBeforeAnswer
    Wscript.Echo "Send enabled: " & objFaxDevice.SendEnabled
    Wscript.Echo "Sending now: " & objFaxDevice.SendingNow
    Wscript.Echo "TSID: " & objFaxDevice.TSID
Next

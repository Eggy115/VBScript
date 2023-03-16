


' List Fax Server Device Providers


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objDeviceProviders = objFaxServer.GetDeviceProviders

For Each objFaxDeviceProvider in objDeviceProviders
    Wscript.Echo "Debug: " & objFaxDeviceProvider.Debug
    Wscript.Echo "Friendly name: " & objFaxDeviceProvider.FriendlyName
    Wscript.Echo "Image name: " & objFaxDeviceProvider.ImageName
    Wscript.Echo "Initialization error code: " & _
        objFaxDeviceProvider.InitErrorCode
    Wscript.Echo "Major build: " & objFaxDeviceProvider.MajorBuild
    Wscript.Echo "Minor build: " & objFaxDeviceProvider.MinorBuild
    Wscript.Echo "Major version: " & objFaxDeviceProvider.MajorVersion
    Wscript.Echo "Minor version: " & objFaxDeviceProvider.MinorVersion
    Wscript.Echo "Status: " & objFaxDeviceProvider.Status
    Wscript.Echo "TAPI provider name: " & objFaxDeviceProvider.TAPIProviderName
    Wscript.Echo "Unique name: " & objFaxDeviceProvider.UniqueName
Next

' List Fax Server Event Logging Options


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Set objFaxLoggingOptions = objFaxServer.LoggingOptions

Set objFaxEventLogging = objFaxLoggingOptions.EventLogging
Wscript.Echo "General events level: " & _
    objFaxEventLogging.GeneralEventsLevel
Wscript.Echo "Inbound events level: " & _
    objFaxEventLogging.InboundEventsLevel
Wscript.Echo "Initialization events level: " & _
    objFaxEventLogging.InitEventsLevel
Wscript.Echo "Outbound events level: " & _
    objFaxEventLogging.OutboundEventsLevel

' List Exchange Server Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery("Select * from Exchange_Server")

For Each objItem in colItems
    Wscript.Echo "Administrative group: " & _
        objItem.AdministrativeGroup
    Wscript.Echo "Administrative note: " & objItem.AdministrativeNote
    Wscript.Echo "Creation time: " & objItem.CreationTime
    Wscript.Echo "Distinguished name: " & objItem.DN
    Wscript.Echo "Exchange version: " & objItem.ExchangeVersion
    Wscript.Echo "Fully-qualified domain name: " & objItem.FQDN
    Wscript.Echo "GUID: " & objItem.GUID
    Wscript.Echo "Is front-end server: " & objItem.IsFrontEndServer
    Wscript.Echo "Last modification time: " & _
        objItem.LastModificationTime
    Wscript.Echo "Message tracking enabled: " & _
        objItem.MessageTrackingEnabled
    Wscript.Echo "Message tracking log file lifetime: " & _
        objItem.MessageTrackingLogFileLifetime
    Wscript.Echo "Message tracking log file path: " & _
        objItem.MessageTrackingLogFilePath
    Wscript.Echo "Monitoring enabled: " & objItem.MonitoringEnabled
    Wscript.Echo "MTA data path: " & objItem.MTADataPath
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Routing group: " & objItem.RoutingGroup
    Wscript.Echo "Subject logging enabled: " & _
        objItem.SubjectLoggingEnabled
    Wscript.Echo "Type: " & objItem.Type
    Wscript.Echo
Next

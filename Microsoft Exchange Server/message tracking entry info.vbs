
' List Exchange Message Tracking Entry Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_MessageTrackingEntry")

For Each objItem in colItems
    Wscript.Echo "Attempted partner server: " &  _
        objItem.AttemtpedPartnerServer
    Wscript.Echo "Client IP: " & objItem.ClientIP
    Wscript.Echo "Client name: " & objItem.ClientName
    Wscript.Echo "Cost: " & objItem.Cost
    Wscript.Echo "Delivery time: " & objItem.DeliveryTime
    Wscript.Echo "Encrypted: " & objItem.Encrypted
    Wscript.Echo "Entry type: " & objItem.EntryType
    Wscript.Echo "Expansion DL: " & objItem.ExpansionDL
    Wscript.Echo "Key ID: " & objItem.KeyID
    Wscript.Echo "Linked message ID: " &  _
        objItem.LinkedMessageID
    Wscript.Echo "Message ID: " & objItem.MessageID
    Wscript.Echo "Origination time: " & objItem.OriginationTime
    Wscript.Echo "Partner server: " & objItem.PartnerServer
    Wscript.Echo "Priority: " & objItem.Priority
    Wscript.Echo "Recipient address: " &  _
        objItem.RecipientAddress
    Wscript.Echo "Recipient count: " & objItem.RecipientCount
    Wscript.Echo "Recipient status: " & objItem.RecipientStatus
    Wscript.Echo "Sender address: " & objItem.SenderAddress
    Wscript.Echo "Server IP: " & objItem.ServerIP
    Wscript.Echo "Server name: " & objItem.ServerName
    Wscript.Echo "Size: " & objItem.Size
    Wscript.Echo "Subject: " & objItem.Subject
    Wscript.Echo "Subject ID: " & objItem.SubjectID
    Wscript.Echo "Time logged: " & objItem.TimeLogged
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo
Next

' List Exchange Mailbox Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_Mailbox")

For Each objItem in colItems
    Wscript.Echo "Associated content count: " & _
        objItem.AssocContentCount
    Wscript.Echo "Date discovered absent in directory service: " & _
        objItem.DateDiscoveredAbsentInDS
    Wscript.Echo "Delete messages size extended: " & _
        objItem.DeletedMessageSizeExtended
    Wscript.Echo "Last logged-on user account: " & _
        objItem.LastLoggedOnUserAccount
    Wscript.Echo "Last logoff time: " & objItem.LastLogoffTime
    Wscript.Echo "Last logon time: " & objItem.LastLogonTime
    Wscript.Echo "Legacy distinguished name: " & objItem.LegacyDN
    Wscript.Echo "Mailbox display name: " & _
        objItem.MailboxDisplayName
    Wscript.Echo "Mailbox GUID: " & objItem.MailboxGUID
    Wscript.Echo "Server name: " & objItem.ServerName
    Wscript.Echo "Size: " & objItem.Size
    Wscript.Echo "Storage group name: " & _
        objItem.StorageGroupName
    Wscript.Echo "Storage limit information: " & _
        objItem.StorageLimitInfo
    Wscript.Echo "Store name: " & objItem.StoreName
    Wscript.Echo "Total items: " & objItem.TotalItems
    Wscript.Echo
Next

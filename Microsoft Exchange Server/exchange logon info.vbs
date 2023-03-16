

' List Exchange Logon Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery("Select * from Exchange_Logon")

For Each objItem in colItems
    Wscript.Echo "Client version: " & objItem.ClientVersion
    Wscript.Echo "Code page ID: " & objItem.CodePageID
    Wscript.Echo "Folder operations rate: " & _
        objItem.FolderOperationsRate
    Wscript.Echo "Host addess: " & objItem.HostAddress
    Wscript.Echo "Last operation time: " & _
        objItem.LastOperationTime
    Wscript.Echo "Locale ID: " & objItem.LocaleID
    Wscript.Echo "Logged-on user account: " & _
        objItem.LoggedOnUserAccount
    Wscript.Echo "Logged-on user's malibx legacy distinguished name: " _
        & objItem.LoggedOnUsersMailboxLegacyDN
    Wscript.Echo "Logon time: " & objItem.LogonTime
    Wscript.Echo "Mailbox display name: " & _
        objItem.MailboxDisplayName
    Wscript.Echo "Mailbox legacy distinguished name: " & _
        objItem.MailboxLegacyDN
    Wscript.Echo "Messaging operation count: " & _
        objItem.MessagingOperationRate
    Wscript.Echo "Open attachment count: " & _
        objItem.OpenAttachmentCount
    Wscript.Echo "Open folder count: " & objItem.OpenFolderCount
    Wscript.Echo "Open message count: " & objItem.OpenMessageCount
    Wscript.Echo "Other operation rate: " & _
        objItem.OtherOperationRate
    Wscript.Echo "Progress operation rate: " & _
        objItem.ProgressOperationRate
    Wscript.Echo "Row ID: " & objItem.RowID
    Wscript.Echo "Server name: " & objItem.ServerName
    Wscript.Echo "Storage group name: " & objItem.StorageGroupName
    Wscript.Echo "Store name: " & objItem.StoreName
    Wscript.Echo "Store type: " & objItem.StoreType
    Wscript.Echo "Stream operation rate: " & _
        objItem.StreamOperationRate
    Wscript.Echo "Table operation rate: " & _
        objItem.TableOperationRate
    Wscript.Echo "Total operation rate: " & _
        objItem.TotalOperationRate
    Wscript.Echo "Transfer operation rate: " & _
        objItem.TransferOperationRate
    Wscript.Echo
Next

' List Exchange Public Folder Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_PublicFolder")

For Each objItem in colItems
    Wscript.Echo "Address book name: " & objItem.AddressBookName
    Wscript.Echo "Administrative note: " & _
        objItem.AdministrativeNote
    Wscript.Echo "Administrative security descriptor: " & _
        objItem.AdminSecurityDescriptor
    Wscript.Echo "Active Directory proxy path: " & _
        objItem.ADProxyPath
    Wscript.Echo "Associated messgae count: " & _
        objItem.AssociatedMessageCount
    Wscript.Echo "Attachment count: " & objItem.AttachmentCount
    Wscript.Echo "Categorization count: " & _
        objItem.CategorizationCount
    Wscript.Echo "Comment: " & objItem.Comment
    Wscript.Echo "Contact count: " & objItem.ContactCount
    Wscript.Echo "Contains rules: " & objItem.ContainsRules
    Wscript.Echo "Creation time: " & objItem.CreationTime
    Wscript.Echo "Deleted item lifetime: " & _
        objItem.DeletedItemLifetime
    Wscript.Echo "Folder tree: " & objItem.FolderTree
    Wscript.Echo "Friendly URL: " & objItem.FriendlyURL
    Wscript.Echo "Has children: " & objItem.HasChildren
    Wscript.Echo "Has local replica: " & objItem.HasLocalReplica
    Wscript.Echo "Is mail enabled: " & objItem.IsMailEnabled
    Wscript.Echo "Is normal folder: " & objItem.IsNormalFolder
    Wscript.Echo "Is search folder: " & objItem.IsSearchFolder
    Wscript.Echo "Is secure in site: " & objItem.IsSecureInSite
    Wscript.Echo "Last access time: " & objItem.LastAccessTime
    Wscript.Echo "Last modification time: " & _
        objItem.LastModificationTime
    Wscript.Echo "Maximum item size: " & objItem.MaximumItemSize
    Wscript.Echo "Message count: " & objItem.MessageCount
    Wscript.Echo "Message with attachments count: " & _
        objItem.MessageWithAttachmentsCount
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Normal message size: " & _
        objItem.NormalMessageSize
    Wscript.Echo "owner count: " & objItem.OwnerCount
    Wscript.Echo "Parent friendly URL: " & _
        objItem.ParentFriendlyURL
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "Prohibit post limit: " & _
        objItem.ProhibitPostLimit
    Wscript.Echo "Publish in address book: " & _
        objItem.PublishInAddressBook
    Wscript.Echo "Recipiejt count on associated messages: " & _
        objItem.RecipientCountOnAssociatedMessages
    Wscript.Echo "Recipient count on normal messages: " & _
        objItem.RecipientCountOnNormalMessages
    Wscript.Echo "Replica age limit: " & objItem.ReplicaAgeLimit
    Wscript.Echo "Replica list: " & objItem.ReplicaList
    Wscript.Echo "Replication message priority: " & _
        objItem.ReplicationMessagePriority
    Wscript.Echo "Replication schedule: " & _
        objItem.ReplicationSchedule
    Wscript.Echo "Replication style: " & objItem.ReplicationStyle
    Wscript.Echo "Replication count: " & objItem.RestrictionCount
    Wscript.Echo "Security descriptor: " & _
        objItem.SecurityDescriptor
    Wscript.Echo "Storage limit style: " & objItem.StorageLimitStyle
    Wscript.Echo "Target address: " & objItem.TargetAddress
    Wscript.Echo "Total message size: " & objItem.TotalMessageSize
    Wscript.Echo "URL: " & objItem.URL
    Wscript.Echo "Use public store age limits: " & _
        objItem.UsePublicStoreAgeLimits
    Wscript.Echo "Use public store deleted item lifetime: " & _
        objItem.UsePublicStoreDeletedItemLifetime
    Wscript.Echo "Warning limit: " & objItem.WarningLimit
    Wscript.Echo
Next

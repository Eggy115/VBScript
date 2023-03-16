' List the Codec Files on a Computer


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_CodecFile")

For Each objItem in colItems
    Wscript.Echo "Access Mask: " & objItem.AccessMask
    Wscript.Echo "Archive: " & objItem.Archive
    Wscript.Echo "Caption: " & objItem.Caption
    strCreationDate = WMIDateStringToDate(objItem.CreationDate)
    Wscript.Echo "Creation Date: " & strCreationdate
    Wscript.Echo "Drive: " & objItem.Drive
    Wscript.Echo "Eight Dot Three File Name: " & _
        objItem.EightDotThreeFileName
    Wscript.Echo "Extension: " & objItem.Extension
    Wscript.Echo "File Name: " & objItem.FileName
    Wscript.Echo "File Size: " & objItem.FileSize
    Wscript.Echo "File Type: " & objItem.FileType
    Wscript.Echo "File System Name: " & objItem.FSName
    Wscript.Echo "Group: " & objItem.Group
    Wscript.Echo "Hidden: " & objItem.Hidden
    strInstallDate = WMIDateStringToDate(objItem.InstallDate)
    Wscript.Echo "Last Accessed: " & strLastAccessed
    strLastModified = WMIDateStringToDate(objItem.LastModified)
    Wscript.Echo "Last Modified: " & strLastModified
    Wscript.Echo "Manufacturer: " & objItem.Manufacturer
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "Version: " & objItem.Version
Next
 
Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
        Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
            & " " & Mid (dtmDate, 9, 2) & ":" & _
                Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate, _
                    13, 2))
End Function

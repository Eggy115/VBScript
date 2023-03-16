' List Shortcuts on a Computer


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_ShortcutFile")

For Each objItem in colItems
    strCreationDate = WMIDateStringToDate(objItem.CreationDate)
    Wscript.Echo "Creation Date: " & strCreationDate
    Wscript.Echo "Drive: " & objItem.Drive
    Wscript.Echo "Eight Dot Three File Name: " & _
        objItem.EightDotThreeFileName
    Wscript.Echo "Extension: " & objItem.Extension
    Wscript.Echo "File Name: " & objItem.FileName
    Wscript.Echo "File Size: " & objItem.FileSize
    Wscript.Echo "File Type: " & objItem.FileType
    Wscript.Echo "File System Name: " & objItem.FSName
    Wscript.Echo "Hidden: " & objItem.Hidden
    strLastAccessed = WMIDateStringToDate(objItem.LastAccessed)
    Wscript.Echo "Last Accessed: " & strLastAccessed
    strLastModified = WMIDateStringToDate(objItem.LastModified)
    Wscript.Echo "Last Modified: " & strLastModified
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Path: " & objItem.Path
    Wscript.Echo "Target: " & objItem.Target
Next
 
Function WMIDateStringToDate(dtmDate)
    WMIDateStringToDate = CDate(Mid(dtmDate, 5, 2) & "/" & _
        Mid(dtmDate, 7, 2) & "/" & Left(dtmDate, 4) _
            & " " & Mid (dtmDate, 9, 2) & ":" & _
                Mid(dtmDate, 11, 2) & ":" & Mid(dtmDate, _
                    13, 2))
End Function

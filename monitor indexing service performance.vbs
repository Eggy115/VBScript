' Monitor Indexing Service Performance


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService," & _
    "Win32_PerfFormattedData_ContentIndex_IndexingService").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Deferred for Indexing: " & objItem.Deferredforindexing
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Files to be Indexed: " & objItem.Filestobeindexed
        Wscript.Echo "Index Size in Megabytes: " & objItem.IndexsizeMB
        Wscript.Echo "Merge Progress: " & objItem.Mergeprogress
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Number of Documents Indexed: " & _
            objItem.Numberdocumentsindexed
        Wscript.Echo "Running Queries: " & objItem.Runningqueries
        Wscript.Echo "Saved Indexes: " & objItem.Savedindexes
        Wscript.Echo "Total Number of Documents: " & _
            objItem.TotalNumberdocuments
        Wscript.Echo "Total Number of Queries: " & objItem.TotalNumberofqueries
        Wscript.Echo "Unique Keys: " & objItem.Uniquekeys
        Wscript.Echo "Word Lists: " & objItem.Wordlists
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

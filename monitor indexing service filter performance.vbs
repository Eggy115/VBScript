' Monitor Indexing Service Filter Performance


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum(objWMIService, " & _
    "Win32_PerfFormattedData_ContentFilter_IndexingServiceFilter").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Binding Time in Milliseconds: " & _
            objItem.Bindingtimemsec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Indexing Speed, Megabytes Per Hour: " & _
            objItem.IndexingspeedMBPerhr
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Total Indexing Speed, Megabytes Per Hour: " & _
            objItem.TotalindexingspeedMBPerhr
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

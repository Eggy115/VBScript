' Monitor Cache Performance


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfOS_Cache").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Asynchronous Copy Reads Per Second: " & _
            objItem.AsyncCopyReadsPersec
        Wscript.Echo "Asynchronous Data Maps Per Second: " & _
            objItem.AsyncDataMapsPersec
        Wscript.Echo "Asynchronous Fast Reads Per Second: " & _
            objItem.AsyncFastReadsPersec
        Wscript.Echo "Asynchronous MDL Reads Per Second: " & _
            objItem.AsyncMDLReadsPersec
        Wscript.Echo "Asynchronous Pin Reads Per Second: " & _
            objItem.AsyncPinReadsPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Copy Read Hits Percent: " & objItem.CopyReadHitsPercent
        Wscript.Echo "Copy Reads Per Second: " & objItem.CopyReadsPersec
        Wscript.Echo "Data Flushes Per Second: " & objItem.DataFlushesPersec
        Wscript.Echo "Data Flush Pages Per Second: " & _
            objItem.DataFlushPagesPersec
        Wscript.Echo "Data Map Hits Percent: " & objItem.DataMapHitsPercent
        Wscript.Echo "Data Map Pins Per Second: " & objItem.DataMapPinsPersec
        Wscript.Echo "Data Maps Per Second: " & objItem.DataMapsPersec
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Fast Read Not Possibles Per Second: " & _
            objItem.FastReadNotPossiblesPersec
        Wscript.Echo "Fast Read Resource Misses Per Second: " & _
            objItem.FastReadResourceMissesPersec
        Wscript.Echo "Fast Reads Per Second: " & objItem.FastReadsPersec
        Wscript.Echo "Lazy Write Flushes Per Second: " & _
            objItem.LazyWriteFlushesPersec
        Wscript.Echo "Lazy Write Pages Per Second: " & _
            objItem.LazyWritePagesPersec
        Wscript.Echo "MDL Read Hits Percent: " & objItem.MDLReadHitsPercent
        Wscript.Echo "MDL Reads Per Second: " & objItem.MDLReadsPersec
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Pin Read Hits Percent: " & objItem.PinReadHitsPercent
        Wscript.Echo "Pin Reads Per Second: " & objItem.PinReadsPersec
        Wscript.Echo "Read Aheads Per Second: " & objItem.ReadAheadsPersec
        Wscript.Echo "Synchronous Copy Reads Per Second: " & _
            objItem.SyncCopyReadsPersec
        Wscript.Echo "Synchronous Data Maps Per Second: " & _
            objItem.SyncDataMapsPersec
        Wscript.Echo "Synchronous Fast Reads Per Second: " & _
            objItem.SyncFastReadsPersec
        Wscript.Echo "Synchronous MDL Reads Per Second: " & _
            objItem.SyncMDLReadsPersec
        Wscript.Echo "Synchronous Pin Reads Per Second: " & _
            objItem.SyncPinReadsPersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next



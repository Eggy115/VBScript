

' Monitor Memory Performance


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
Set colItems = objRefresher.AddEnum _
    (objWMIService, "Win32_PerfFormattedData_PerfOS_Memory").objectSet
objRefresher.Refresh

For i = 1 to 5
    For Each objItem in colItems
        Wscript.Echo "Available Bytes: " & objItem.AvailableBytes
        Wscript.Echo "Available Kilobytes: " & objItem.AvailableKBytes
        Wscript.Echo "Available Megabytes: " & objItem.AvailableMBytes
        Wscript.Echo "Cache Bytes: " & objItem.CacheBytes
        Wscript.Echo "Cache Bytes Peak: " & objItem.CacheBytesPeak
        Wscript.Echo "Cache Faults Per Second: " & objItem.CacheFaultsPersec
        Wscript.Echo "Caption: " & objItem.Caption
        Wscript.Echo "Commit Limit: " & objItem.CommitLimit
        Wscript.Echo "Committed Bytes: " & objItem.CommittedBytes
        Wscript.Echo "Demand Zero Faults Per Second: " & _
            objItem.DemandZeroFaultsPersec
        Wscript.Echo "Description: " & objItem.Description
        Wscript.Echo "Free System Page Table Entries: " & _
            objItem.FreeSystemPageTableEntries
        Wscript.Echo "Name: " & objItem.Name
        Wscript.Echo "Page Faults Per Second: " & objItem.PageFaultsPersec
        Wscript.Echo "Page Reads Per Second: " & objItem.PageReadsPersec
        Wscript.Echo "Pages Input Per Second: " & objItem.PagesInputPersec
        Wscript.Echo "Pages Output Per Second: " & objItem.PagesOutputPersec
        Wscript.Echo "Pages Per Second: " & objItem.PagesPersec
        Wscript.Echo "Page Writes Per Second: " & objItem.PageWritesPersec
        Wscript.Echo "Percent Committed Bytes In Use: " & _
            objItem.PercentCommittedBytesInUse
        Wscript.Echo "Pool Nonpaged Allocations: " & objItem.PoolNonpagedAllocs
        Wscript.Echo "Pool Nonpaged Bytes: " & objItem.PoolNonpagedBytes
        Wscript.Echo "Pool Paged Allocations: " & objItem.PoolPagedAllocs
        Wscript.Echo "Pool Paged Bytes: " & objItem.PoolPagedBytes
        Wscript.Echo "Pool Paged Resident Bytes: " & _
            objItem.PoolPagedResidentBytes
        Wscript.Echo "System Cache Resident Bytes: " & _
            objItem.SystemCacheResidentBytes
        Wscript.Echo "System Code Resident Bytes: " & _
            objItem.SystemCodeResidentBytes
        Wscript.Echo "System Code Total Bytes: " & objItem.SystemCodeTotalBytes
        Wscript.Echo "System Driver Resident Bytes: " & _
            objItem.SystemDriverResidentBytes
        Wscript.Echo "System Driver Total Bytes: " & _
            objItem.SystemDriverTotalBytes
        Wscript.Echo "Transition Faults Per Second: " & _
            objItem.TransitionFaultsPersec
        Wscript.Echo "Write Copies Per Second: " & objItem.WriteCopiesPersec
        Wscript.Sleep 2000
        objRefresher.Refresh
    Next
Next

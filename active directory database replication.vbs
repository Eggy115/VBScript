
' List Active Directory Database Replication Partners


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & _
        strComputer & "\root\MicrosoftActiveDirectory")

Set colReplicationOperations = objWMIService.ExecQuery _
    ("Select * from MSAD_ReplNeighbor")

For each objReplicationJob in colReplicationOperations 
    Wscript.Echo "Domain: " & objReplicationJob.Domain
    Wscript.Echo "Naming context DN: " & objReplicationJob.NamingContextDN
    Wscript.Echo "Source DSA DN: " & objReplicationJob.SourceDsaDN
    Wscript.Echo "Last synch result: " & objReplicationJob.LastSyncResult
    Wscript.Echo "Number of consecutive synchronization failures: " & _
        objReplicationJob.NumConsecutiveSyncFailures
Next

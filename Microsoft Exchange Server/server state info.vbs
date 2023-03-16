' List Exchange Server State Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" &  _
        strComputer & "\root\cimv2\Applications\Exchange")

Set colItems = objWMIService.ExecQuery _
    ("Select * from ExchangeServerState")

For Each objItem in colItems
    Wscript.Echo "Cluster state: " & objItem.ClusterState
    Wscript.Echo "Cluster state string: " & _
        objItem.ClusterStateString
    Wscript.Echo "CPU state: " & objItem.CPUState
    Wscript.Echo "CPU state string: " & objItem.CPUStateString
    Wscript.Echo "Disks state: " & objItem.DisksState
    Wscript.Echo "Diskss state string: " & objItem.DisksStateString
    Wscript.Echo "Distinguished name: " & objItem.DN
    Wscript.Echo "Group dsitinguihsed name: " & objItem.GroupDN
    Wscript.Echo "Group GUID: " & objItem.GroupGUID
    Wscript.Echo "GUID: " & objItem.GUID
    Wscript.Echo "Memory state: " & objItem.MemoryState
    Wscript.Echo "Memory state string: " & _
        objItem.MemoryStateString
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Queues state: " & objItem.QueuesState
    Wscript.Echo "Queues state string: " & _
        objItem.QueuesStateString
    Wscript.Echo "Server maintenance: " & _
        objItem.ServerMaintenance
    Wscript.Echo "Server state: " & objItem.ServerState
    Wscript.Echo "Server state string: " & _
        objItem.ServerStateString
    Wscript.Echo "Services state: " & objItem.ServicesState
    Wscript.Echo "Services state string: " & _
        objItem.ServicesStateString
    Wscript.Echo "Unreachable: " & objItem.Unreachable
    Wscript.Echo "Version: " & objItem.Version
    Wscript.Echo
Next

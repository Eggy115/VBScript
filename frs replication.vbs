' Monitor FRS Replication


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colFRSSet = objWMIService.ExecQuery _   
 ("Select * from Win32_PerfFormattedData_FileReplicaConn_FileReplicaConn")

For Each objFRSInstance in colFRSSet 
    Wscript.Echo "Remote change orders received: " & _
        objFRSInstance.RemoteChangeOrdersReceived
    Wscript.Echo "Remote change orders sent: " & _
        objFRSInstance.RemoteChangeOrdersSent
    Wscript.Echo "Packets sent: " & objFRSInstance.PacketsSent
Next

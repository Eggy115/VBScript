' List the Results of the Last System Restore


strComputer = "."
 
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")

Set objItem = objWMIService.Get("SystemRestore")
errResults = objItem.GetLastRestoreStatus()
 
Select Case errResults
    Case 0 strRestoreStatus = "The last restore failed."
    Case 1 strRestoreStatus = "The last restore was successful."
    Case 2 strRestoreStatus = "The last restore was interrupted."
End Select
 
Wscript.Echo strRestoreStatus



' List the State of the Indexing Service


On Error Resume Next

Set objAdminIS = CreateObject("Microsoft.ISAdm")
Wscript.Echo "Is running: " & objAdminIS.IsRunning
Wscript.Echo "Is paused: " & objAdminIS.IsPaused
Wscript.Echo "Computer name: " & objAdminIS.MachineName

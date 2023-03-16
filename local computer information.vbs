
' List Local Computer Information


Set objComputer = CreateObject("Shell.LocalMachine")

Wscript.Echo "Computer name: " & objComputer.MachineName
Wscript.Echo "Shutdown allowed: " & objComputer.IsShutdownAllowed
Wscript.Echo "Friendly UI enabled: " & objComputer.IsFriendlyUIEnabled
Wscript.Echo "Guest access mode: " & objComputer.IsGuestAccessMode
Wscript.Echo "Guest account enabled: " & _
    objComputer.IsGuestEnabled(0)
Wscript.Echo "Multiple users enabled: " & _
    objComputer.IsMultipleUsersEnabled
Wscript.Echo "Offline files enabled: " & _
    objComputer.IsOfflineFilesEnabled
Wscript.Echo "Remote connections enabled: " & _
    objComputer.IsRemoteConnectionsEnabled
Wscript.Echo "Undock enabled: " & objComputer.IsUndockEnabled

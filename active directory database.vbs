' Install Active Directory Database Performance Counters


Set WshShell = WScript.CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objShell = CreateObject("Shell.Application")
objFSO.CreateFolder ("C:\Performance")
Set objCopyFile = objFSO.GetFile("C:\windows\system32\esentprf.dll ")
objCopyFile.Copy ("C:\performance\esentprf.dll ") 

WshShell.RegWrite _
    "HKLM\System\CurrentControlSet\Services\Esent\Performance\Open", _
        "OpenPerformanceData", "REG_SZ"
WshShell.RegWrite _
    "HKLM\System\CurrentControlSet\Services\Esent\Performance\Collect", _
        "CollectPerformanceData", "REG_SZ"
WshShell.RegWrite _
    "HKLM\System\CurrentControlSet\Services\Esent\Performance\Close", _
        "ClosePerformanceData", "REG_SZ"
WshShell.RegWrite _
    "HKLM\System\CurrentControlSet\Services\Esent\Performance\Library", _
        "C:\Performance\Esentprf.dll", "REG_SZ"
strCommandText = "%comspec% /c lodctr.exe c:\windows\system32\esentprf.ini" 
WshShell.Run strCommandText

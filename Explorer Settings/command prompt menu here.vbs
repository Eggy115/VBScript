' Add "Command Prompt Here" to Windows Explorer


Set objShell = CreateObject("WScript.Shell")
 
objShell.RegWrite "HKCR\Folder\Shell\MenuText\Command\", _
    "cmd.exe /k cd " & chr(34) & "%1" & chr(34)
objShell.RegWrite "HKCR\Folder\Shell\MenuText\", "Command Prompt Here"

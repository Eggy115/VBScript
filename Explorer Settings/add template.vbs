' Add a Template to the Windows Explorer New Menu


Set objShell = WScript.CreateObject("WScript.Shell")
objShell.RegWrite "HKCR\.VBS\ShellNew\FileName","template.vbs"

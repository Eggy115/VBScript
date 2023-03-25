Const CONST_HIDE_WINDOW = 0
Dim oShell, objArgs
Dim strCmd
Dim I

strCmd = ""
Set objArgs = WScript.Arguments
For I = 0 To objArgs.Count - 1
  strCmd = strCmd & " " & objArgs(I)
Next

' Run Command with hidden style
Set oShell = WScript.CreateObject("WScript.shell")
oShell.Run "CMD.exe /c " & strCmd, CONST_HIDE_WINDOW, TRUE
Set oShell = Nothing

WScript.Quit(0)

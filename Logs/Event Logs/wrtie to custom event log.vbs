' Write to a Custom Event Log Using EventCreate


Set WshShell = WScript.CreateObject("WScript.Shell")

strCommand = "eventcreate /T Error /ID 100 /L Scripts /D " & _
    Chr(34) & "Test event." & Chr(34)
WshShell.Run strcommand

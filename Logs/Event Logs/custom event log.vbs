
' Create a Custom Event Log


Const NO_VALUE = Empty

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.RegWrite _
    "HKLM\System\CurrentControlSet\Services\EventLog\Scripts\", NO_VALUE

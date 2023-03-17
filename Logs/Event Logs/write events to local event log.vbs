


' Write Events to the Local Event Log


Const EVENT_SUCCESS = 0

Set objShell = Wscript.CreateObject("Wscript.Shell")

objShell.LogEvent EVENT_SUCCESS, _
    "Payroll application successfully installed."

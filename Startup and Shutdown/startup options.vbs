' List Computer Startup Options


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colStartupCommands = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")

For Each objStartupCommand in colStartupCommands
    Wscript.Echo "Reset Boot Enabled: " & _
        objStartupCommand.AutomaticResetBootOption
    Wscript.Echo "Reset Boot Possible: " & _
        objStartupCommand.AutomaticResetCapability
    Wscript.Echo "Boot State: " & objStartupCommand.BootupState
    Wscript.Echo "Startup Delay: " & objStartupCommand.SystemStartupDelay
    For i = 0 to Ubound(objStartupCommand.SystemStartupOptions)
        Wscript.Echo "Startup Options: " & _
            objStartupCommand.SystemStartupOptions(i)
    Next
    Wscript.Echo "Startup Setting: " & _
        objStartupCommand.SystemStartupSetting
Next

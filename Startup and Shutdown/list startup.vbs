' List Computer Startup Commands


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colStartupCommands = objWMIService.ExecQuery _
    ("Select * from Win32_StartupCommand")

For Each objStartupCommand in colStartupCommands
    Wscript.Echo "Command: " & objStartupCommand.Command
    Wscript.Echo "Description: " & objStartupCommand.Description
    Wscript.Echo "Location: " & objStartupCommand.Location
    Wscript.Echo "Name: " & objStartupCommand.Name
    Wscript.Echo "Setting ID: " & objStartupCommand.SettingID
    Wscript.Echo "User: " & objStartupCommand.User
Next


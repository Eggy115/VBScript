' List the Boot Configuration Properties of a Computer


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_BootConfiguration")

For Each objItem in colItems
    Wscript.Echo "Boot Directory: " & objItem.BootDirectory
    Wscript.Echo "Configuration Path: " & objItem.ConfigurationPath
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Last Drive: " & objItem.LastDrive
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Scratch Directory: " & objItem.ScratchDirectory
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Temp Directory: " & objItem.TempDirectory
Next


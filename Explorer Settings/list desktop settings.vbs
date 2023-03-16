

' List Desktop Settings


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery("Select * from Win32_Desktop")

For Each objItem in colItems
    Wscript.Echo "Border Width: " & objItem.BorderWidth
    Wscript.Echo "Caption: " & objItem.Caption
    Wscript.Echo "Cool Switch: " & objItem.CoolSwitch
    Wscript.Echo "Cursor Blink Rate: " & objItem.CursorBlinkRate
    Wscript.Echo "Description: " & objItem.Description
    Wscript.Echo "Drag Full Windows: " & objItem.DragFullWindows
    Wscript.Echo "Grid Granularity: " & objItem.GridGranularity
    Wscript.Echo "Icon Spacing: " & objItem.IconSpacing
    Wscript.Echo "Icon Title Face Name: " & objItem.IconTitleFaceName
    Wscript.Echo "Icon Title Size: " & objItem.IconTitleSize
    Wscript.Echo "Icon Title Wrap: " & objItem.IconTitleWrap
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Pattern: " & objItem.Pattern
    Wscript.Echo "Screen Saver Active: " & objItem.ScreenSaverActive
    Wscript.Echo "Screen Saver Executable: " & _
        objItem.ScreenSaverExecutable
    Wscript.Echo "Screen Saver Secure: " & objItem.ScreenSaverSecure
    Wscript.Echo "Screen Saver Timeout: " & objItem.ScreenSaverTimeout
    Wscript.Echo "Setting ID: " & objItem.SettingID
    Wscript.Echo "Wallpaper: " & objItem.Wallpaper
    Wscript.Echo "Wallpaper Stretched: " & objItem.WallpaperStretched
    Wscript.Echo "Wallpaper Tiled: " & objItem.WallpaperTiled
Next

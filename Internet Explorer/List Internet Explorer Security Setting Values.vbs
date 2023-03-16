' List Internet Explorer Security Setting Values


On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
strEntry = "1400"

Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "Zones\1"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, strEntry, dwValue

Select Case dwValue
    Case 0 strSetting = "Enabled"
    Case 1 strSetting = "Prompt"
    case 3 strSetting = "Disabled"
End Select

Wscript.Echo "Allow scripting: " & strSetting

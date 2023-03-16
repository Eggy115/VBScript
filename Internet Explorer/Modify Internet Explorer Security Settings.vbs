' Modify Internet Explorer Security Settings


On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001

strComputer = "."

Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "Zones\1"
strEntryName = "1400"
dwvalue = 0
objReg.SetDWORDValue HKEY_CURRENT_USER, strKeyPath, strEntryName,dwValue

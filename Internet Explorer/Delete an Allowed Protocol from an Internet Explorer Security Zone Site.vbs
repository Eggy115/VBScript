' Delete an Allowed Protocol from an Internet Explorer Security Zone Site


On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001

strComputer = "."

Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\ESCDomains\Finance"
strDWORDValueName = "http"

objReg.DeleteValue HKEY_CURRENT_USER,strKeyPath,strDWORDValueName

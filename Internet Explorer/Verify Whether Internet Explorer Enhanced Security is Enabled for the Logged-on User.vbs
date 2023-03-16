'  Verify Whether Internet Explorer Enhanced Security is Enabled for the Logged-on User


On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Internet " _
    & "Settings\ZoneMap"
strValueName = "IEHarden"
objReg.GetDWORDValue HKEY_CURRENT_USER,strKeyPath,strValueName,intHarden
 
If intHarden = 1 Then
    Wscript.Echo "IE hardening is turned on for the current user."
Else
    Wscript.Echo "IE hardening is not turned on for the current user."
End If

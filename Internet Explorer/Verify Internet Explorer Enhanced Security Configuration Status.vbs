' Verify Internet Explorer Enhanced Security Configuration Status


On Error Resume Next

Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Active Setup\Installed Components\" _
    & "{A509B1A7-37EF-4b3f-8CFC-4F3A74704073}"
strValueName = "IsInstalled"
objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,intAdmin
 
strKeyPath = "SOFTWARE\Microsoft\Active Setup\Installed Components\" _
    & "{A509B1A8-37EF-4b3f-8CFC-4F3A74704073}"
strValueName = "IsInstalled"
objReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,intUsers

strConfiguration = intAdmin & intUsers
Select Case strConfiguration
    Case "00"
        Wscript.Echo "The use of Internet Explorer is not restricted on " _
            & "this server."
    Case "01"
        Wscript.Echo "The use of Internet Explorer is restricted for the " _
           & "administrators group on this server. The use of Internet " _
           & "Explorer is not restricted for any other user group."
    Case "10"
        Wscript.Echo "The use of Internet Explorer is not restricted for the" _
            & " administrators group on this server. The use of Internet " _
            & "Explorer is restricted for any other user group."
    Case "11"
        Wscript.Echo "The use of Internet Explorer is restricted for all " _
            & "user groups on this server."
End Select

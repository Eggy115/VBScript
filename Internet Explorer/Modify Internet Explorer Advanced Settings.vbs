' Modify Internet Explorer Advanced Settings


On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001

strComputer = "."
strEntry = "DisplayTrustAlertDlg"

Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Internet Explorer\Main"
objReg.GetDWORDValue HKEY_CURRENT_USER, strKeyPath, strEntry ,dwValue

If dwValue = 1 Then
    Wscript.Echo "Enhanced security dialog box is displayed." 
Else
    Wscript.Echo "Enhanced security dialog box is not displayed." 
End If

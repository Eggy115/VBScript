Const HKCU = &H80000001 'HKEY_CURRENT_USER

sComputer = "."   

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _
            & sComputer & "\root\default:StdRegProv")

sKeyPath = "Control Panel\Desktop\"
sValueName = "TranscodedImageCache"
oReg.GetBinaryValue HKCU, sKeyPath, sValueName, sValue


sContents = ""

For i = 24 To UBound(sValue)
  vByte = sValue(i)
  If vByte <> 0 And vByte <> "" Then
    sContents = sContents & Chr(vByte)
  End If
Next

CreateObject("Wscript.Shell").Run "explorer.exe /select,""" & sContents & """"
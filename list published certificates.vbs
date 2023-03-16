
' List Published Certificates for a User Account


On Error Resume Next

Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D
Const ForWriting = 2
Const WshRunning = 0
 
Set objUser = GetObject _
    ("GC://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
objUser.GetInfoEx Array("userCertificate"), 0
arrUserCertificates = objUser.GetEx("userCertificate")
 
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "No assigned certificates"
    WScript.Quit
Else
    Set objShell = CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strPath = "." 
    intFileCounter = 0
 
    For Each arrUserCertificate in arrUserCertificates
        strFileName = "file" & intFileCounter
        strFullName = objFSO.BuildPath(strPath, strFileName)
        Set objFile = objFSO.OpenTextFile(strFullName, ForWriting, True)
        
        For i = 1 To LenB(arrUserCertificate)
            ReDim Preserve arrUserCertificatesChar(i - 1)
            arrUserCertificatesChar(i-1) = _
                Hex(AscB(MidB(arrUserCertificate, i, 3)))
        Next
                
        intCounter=0
        For Each HexVal in arrUserCertificatesChar
            intCounter=intCounter + 1
            If Len(HexVal) = 1 Then 
                objFile.Write(0 & HexVal & " ")
            Else
                objFile.Write(HexVal & " ")
            End If
        Next
        objFile.Close
        Set objFile = Nothing
  
        Set objExecCmd1 = objShell.Exec _
            ("certutil -decodeHex " & strFileName & " " & strFileName & ".cer")
        Do While objExecCmd1.Status = WshRunning
            WScript.Sleep 100
        Loop
        Set objExecCmd1 = Nothing
 
        Set objExecCmd2 = objShell.Exec("certutil " & strFileName & ".cer")
        Set objStdOut = objExecCmd2.StdOut
        Set objExecCmd2 = Nothing
      
        WScript.Echo VbCrLf & "Certificate " & intFileCounter + 1
        While Not objStdOut.AtEndOfStream
            strLine = objStdOut.ReadLine
            If InStr(strLine, "Issuer:") Then
                WScript.Echo Trim(strLine)
                WScript.Echo vbTab & Trim(objStdOut.ReadLine)
            End If
            If InStr(strLine, "Subject:") Then
                Wscript.Echo Trim(strLine)
                WScript.Echo vbTab & Trim(objStdOut.ReadLine)
            End If
            If InStr(strLine, "NotAfter:") Then
                strLine = Trim(strLine)
                WScript.Echo "Expires:"
                Wscript.Echo vbTab & Mid(strLine, 11)
            End If
        Wend
 
        objFSO.DeleteFile(strFullName)
        objFSO.DeleteFile(strPath & "\" & strFileName & ".cer") 
  
        intFileCounter = intFileCounter + 1
    Next
End If

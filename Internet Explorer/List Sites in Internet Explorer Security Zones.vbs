' List Sites in Internet Explorer Security Zones


On Error Resume Next

Const HKEY_CURRENT_USER = &H80000001

strComputer = "."

Set objReg = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}\\" & strComputer & _
        "\root\default:StdRegProv")

strKeyPath = "Software\Microsoft\Windows\CurrentVersion\Internet Settings\" _
    & "ZoneMap\ESCDomains"
objReg.EnumKey HKEY_CURRENT_USER, strKeyPath, arrSubKeys

For Each subkey In arrSubKeys
    strNewPath = strKeyPath & "\" & subkey
    ShowSubkeys
Next

Sub ShowSubkeys
    arrPath = Split(strNewPath, "\")
    intSiteName = Ubound(arrPath)
    strSiteName = arrPath(intSiteName)
    objReg.EnumValues HKEY_CURRENT_USER, strNewPath, arrEntries, arrValueTypes

    If Not IsArray(arrEntries) Then
        arrPath = Split(strNewPath, "\")
        intSiteName = Ubound(arrPath)
        strSiteName = arrPath(intSiteName)
        Wscript.Echo strsitename
            objReg.EnumKey HKEY_CURRENT_USER, strNewPath, arrSubKeys2

        For Each subkey In arrSubKeys2
            strNewPath2 = strNewPath & "\" & subkey
            arrPath = Split(strNewPath2, "\")
            intSiteName = Ubound(arrPath)
            strSiteName = arrPath(intSiteName)
            objReg.EnumValues HKEY_CURRENT_USER, strNewPath2, arrEntries2,_
                arrValueTypes

            For i = 0 to Ubound(arrEntries2)
                objReg.GetDWORDValue HKEY_CURRENT_USER, strNewPath2, _
                    arrEntries2(i),dwValue
            Next

            Select Case dwValue
                Case 0 strZone = "My Computer"
                Case 1 strZone = "Local Intranet zone"
                Case 2 strZone = "Trusted Sites Zone"
                Case 3 strZone = "Internet Zone"
                Case 4 strZone = "Restricted Sites Zone"   
            End Select

            Wscript.Echo vbtab & strSiteName & " -- " & strZone
       Next
    End If

    For i = 0 to Ubound(arrEntries)
        objReg.GetDWORDValue HKEY_CURRENT_USER, strNewPath, _
            arrEntries(i),dwValue
    Next
        
    Select Case dwValue
        Case 0 strZone = "My Computer"
        Case 1 strZone = "Local Intranet zone"
        Case 2 strZone = "Trusted Sites Zone"
        Case 3 strZone = "Internet Zone"
        Case 4 strZone = "Restricted Sites Zone"   
    End Select

    Wscript.Echo strSiteName & " -- " & strZone

End Sub

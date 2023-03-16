
' List Active Directory Connections


strDcRDN   = "cn=atl-dc-01"
strSiteRDN = "cn=Ga-Atl-Sales"
 
Set objRootDSE = GetObject("LDAP://RootDSE")
strConfigurationNC = objRootDSE.Get("configurationNamingContext")
 
strNtdsSettingsPath = "LDAP://cn=NTDS Settings," & strDcRDN & _
    ",cn=Servers," & strSiteRDN & ",cn=Sites," & strConfigurationNC
 
Set objNtdsSettings = GetObject(strNtdsSettingsPath)
 
objNtdsSettings.Filter = Array("nTDSConnection")
 
WScript.Echo strDcRDN & " NTDS Connection Objects" & vbCrLf & _
    String(Len(strDcRDN) + 24, "=")
 
For Each objConnection In objNtdsSettings
    WScript.Echo "Name:      " & objConnection.Name
    WScript.Echo "Enabled:   " & objConnection.enabledConnection
    WScript.Echo "From:      " & Split(objConnection.fromServer, ",")(1)
    WScript.Echo "Options:   " & objConnection.Options
    WScript.Echo "Transport: " & Split(objConnection.transportType, ",")(0)
    WScript.Echo "Naming Contexts"
    WScript.Echo "---------------"
    For Each objDNWithBin In objConnection.GetEx("ms-DS-ReplicatesNCReason")
        Wscript.Echo objDNWithBin.DNString
    Next
    WScript.Echo
Next

' List the Dial-In Property Configuration Settings for a User Account


On Error Resume Next

Const E_ADS_PROPERTY_NOT_FOUND  = &h8000500D

Const FourthOctet = 1
Const ThirdOctet = 256
Const SecondOctet = 65536
Const FirstOctet = 16777216
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
blnMsNPAllowDialin = objUser.Get("msNPAllowDialin")
WScript.Echo "Remote Access Permission (Dial-in or VPN)"
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "Control access through Remote Access Policy"
    Err.Clear
Else
    If blnMsNPAllowDialin = True Then
        WScript.Echo "Allow access (msNPAllowDialin)"
    Else
        WScript.Echo "Deny access (msNPAllowDialin)"
    End If
End If
WScript.Echo 
 
arrMsNPSavedCallingStationID = objUser.GetEx("msNPSavedCallingStationID")
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "No Caller-ID specified."
    Err.Clear
Else
    WScript.Echo "Verify Caller ID (msNPSavedCallingStationID): "
    For Each strValue in arrMsNPSavedCallingStationID
        WScript.echo strValue
    Next
  
    objUser.GetEx "msNPCallingStationID"
    If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
        WScript.Echo "Calling station ID(s) specified but not assigned."
        Err.Clear
    Else
        WScript.echo "Calling station ID(s) assigned."
    End If
  
End If
WScript.Echo 
 
intMsRADIUSServiceType = objUser.Get("msRADIUSServiceType")
WScript.Echo "Callback Options"
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "No Callback"
    Err.Clear
Else
    strMsRADIUSCallbackNumber = objUser.Get("msRADIUSCallbackNumber")
    If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
        WScript.Echo "Set by caller (Routing and Remote Access Service only)"
        Err.Clear
  
    strMsRASSavedCallbackNumber = objUser.Get("msRASSavedCallbackNumber")
    If Err.Number <> E_ADS_PROPERTY_NOT_FOUND Then
        WScript.Echo "Unused value of " & strMsRASSavedCallbackNumber & _
            " appears in the Always Callback to field."
    Else
        Err.Clear
    End If  
Else
    WScript.Echo "Always Callback to: " & _
        strMsRADIUSCallbackNumber & " (msRADIUSCallbackNumber)"
    End If
End If   
WScript.Echo
 
intMsRASSavedFramedIPAddress = objUser.Get("msRASSavedFramedIPAddress")
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "No static IP address assigned."
    Err.Clear
Else
    If sgn(intMsRASSavedFramedIPAddress) = -1 Then
        intIP = intMsRASSavedFramedIPAddress
        WScript.StdOut.Write 256 + (int(intIP/FirstOctet)) & "."
        intFirstRemainder = intIP mod FirstOctet
        WScript.StdOut.Write 256 + (int(intFirstRemainder/SecondOctet)) & "."
        intSecondRemainder = intFirstRemainder mod SecondOctet
        WScript.StdOut.Write 256 + (int(intSecondRemainder/ThirdOctet)) & "."
        intThirdRemainder = intSecondRemainder mod ThirdOctet
        WScript.Echo 256 + (int(intThirdRemainder/FourthOctet))
    Else
        intIP = intMsRASSavedFramedIPAddress
        WScript.StdOut.Write  int(intIP/FirstOctet) & "."
        intFirstRemainder = intIP mod FirstOctet
        WScript.StdOut.Write  int(intFirstRemainder/SecondOctet) & "."
        intSecondRemainder = intFirstRemainder mod SecondOctet
        WScript.StdOut.Write  int(intSecondRemainder/ThirdOctet) & "."
        intThirdRemainder = intSecondRemainder mod ThirdOctet
        WScript.Echo int(intThirdRemainder/FourthOctet)
    End If
    
    objUser.Get "msRADIUSFramedIPAddress"
    If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
        WScript.Echo "Static IP address specified but not assigned."
        Err.Clear
    Else
        WScript.Echo "Static IP Address assigned."
    End If
 
End If
WScript.Echo 
 
arrMsRASSavedFramedRoute = objUser.GetEx("msRASSavedFramedRoute")
If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
    WScript.Echo "No static Routes specified."
    Err.Clear
Else
    WScript.echo "Static Routes (msRASSavedFramedRoute):"
    WScript.Echo vbTab & "CIDR 0.0.0.0 Metric"
    For Each strValue in arrMsRASSavedFramedRoute
        WScript.echo vbTab & strValue
    Next
  
    objUser.GetEx "msRADIUSFramedRoute"
    If Err.Number = E_ADS_PROPERTY_NOT_FOUND Then
        WScript.Echo "Static Routes specified but not assigned."
        Err.Clear
    Else
        WScript.echo "Static Routes assigned."
    End If
End If

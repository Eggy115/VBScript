
' List When a Password Expires


Const SEC_IN_DAY = 86400
Const ADS_UF_DONT_EXPIRE_PASSWD = &h10000
 
Set objUserLDAP = GetObject _
  ("LDAP://CN=myerken,OU=management,DC=fabrikam,DC=com")
intCurrentValue = objUserLDAP.Get("userAccountControl")
 
If intCurrentValue and ADS_UF_DONT_EXPIRE_PASSWD Then
    Wscript.Echo "The password does not expire."
Else
    dtmValue = objUserLDAP.PasswordLastChanged 
    Wscript.Echo "The password was last changed on " & _
        DateValue(dtmValue) & " at " & TimeValue(dtmValue) & VbCrLf & _
            "The difference between when the password was last set" &  _
                "and today is " & int(now - dtmValue) & " days"
    intTimeInterval = int(now - dtmValue)
  
    Set objDomainNT = GetObject("WinNT://fabrikam")
    intMaxPwdAge = objDomainNT.Get("MaxPasswordAge")
    If intMaxPwdAge < 0 Then
        WScript.Echo "The Maximum Password Age is set to 0 in the " & _
            "domain. Therefore, the password does not expire."
    Else
        intMaxPwdAge = (intMaxPwdAge/SEC_IN_DAY)
        Wscript.Echo "The maximum password age is " & intMaxPwdAge & " days"
        If intTimeInterval >= intMaxPwdAge Then
          Wscript.Echo "The password has expired."
        Else
          Wscript.Echo "The password will expire on " & _
              DateValue(dtmValue + intMaxPwdAge) & " (" & _
                  int((dtmValue + intMaxPwdAge) - now) & " days from today" & _
                      ")."
        End If
    End If
End If

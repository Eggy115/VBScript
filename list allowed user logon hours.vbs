' List Allowed User Logon Hours


On Error Resume Next
Dim arrLogonHoursBytes(20)
Dim arrLogonHoursBits(167)
arrDayOfWeek = Array _
    ("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
 
Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
arrLogonHours = objUser.Get("logonHours")
 
For i = 1 To LenB(arrLogonHours)
    arrLogonHoursBytes(i-1) = AscB(MidB(arrLogonHours, i, 1))
    WScript.Echo "MidB returns: " & MidB(arrLogonHours, i, 1)
    WScript.Echo "arrLogonHoursBytes: " & arrLogonHoursBytes(i-1)
    wscript.echo vbcrlf
Next
 
intCounter = 0
intLoopCounter = 0
WScript.echo "Day  Byte 1   Byte 2   Byte 3"
For Each LogonHourByte In arrLogonHoursBytes
    arrLogonHourBits = GetLogonHourBits(LogonHourByte)
 
    If intCounter = 0 Then
        WScript.STDOUT.Write arrDayOfWeek(intLoopCounter) & Space(2)
        intLoopCounter = intLoopCounter + 1
    End If
 
    For Each LogonHourBit In arrLogonHourBits
        WScript.STDOUT.Write LogonHourBit
        intCounter = 1 + intCounter
 
        If intCounter = 8 or intCounter = 16 Then
            Wscript.STDOUT.Write Space(1)
        End If
        
        If intCounter = 24 Then
            WScript.echo vbCr
            intCounter = 0
        End If 
    Next
Next
 
Function GetLogonHourBits(x)
    Dim arrBits(7)
    For i = 7 to 0 Step -1
        If x And 2^i Then
            arrBits(i) = 1
        Else
            arrBits(i) = 0
        End If
    Next
    GetLogonHourBits = arrBits
End Function

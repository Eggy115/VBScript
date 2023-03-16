' Read a Comma Separated Values Log



Const ForReading = 1

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("C:\Windows\System32\DHCP\" & _
 "DhcpSrvLog-Mon.log", ForReading)

Wscript.Echo vbCrLf & "DHCP Records"

Do While objTextFile.AtEndOfStream <> True
  strLine = objtextFile.ReadLine
  If inStr(strLine, ",") Then
    arrDHCPRecord = split(strLine, ",")
    Wscript.Echo vbCrLf & "Event ID: " & arrDHCPRecord(0)
    Wscript.Echo "Date: " & arrDHCPRecord(1)
    Wscript.Echo "Time: " & arrDHCPRecord(2)
    Wscript.Echo "Description: " & arrDHCPRecord(3)
    Wscript.Echo "IP Address: " & arrDHCPRecord(4)
    Wscript.Echo "Host Name: " & arrDHCPRecord(5)
    Wscript.Echo "MAC Address: " & arrDHCPRecord(6)
    i = i + 1
  End If
Loop

Wscript.Echo vbCrLf & "Number of DHCP records read: " & i

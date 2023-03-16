' List the User Logged on to a Remote Computer


strComputer = "atl-ws-o1"
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2") 

Set colComputer = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
 
For Each objComputer in colComputer
    Wscript.Echo "Logged-on user: " & objComputer.UserName
Next

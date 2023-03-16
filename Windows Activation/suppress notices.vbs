' Suppress Windows Activation Notices


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colWPASettings = objWMIService.ExecQuery _
    ("Select * from Win32_WindowsProductActivation")
 
For Each objWPASetting in colWPASettings
    objWPASetting.SetNotification(0)
Next

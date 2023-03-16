' Activate Windows Offline


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colWindowsProducts = objWMIService.ExecQuery _
    ("Select * from Win32_WindowsProductActivation")

For Each objWindowsProduct in colWindowsProducts
    objWindowsProduct.ActivateOffline("1234-1234")
Next

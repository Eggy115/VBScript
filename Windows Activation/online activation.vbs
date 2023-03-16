
' Activate Windows Online


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colWindowsProducts = objWMIService.ExecQuery _
    ("Select * from Win32_WindowsProductActivation")

For Each objWindowsProduct in colWindowsProducts
    objWindowsProduct.ActivateOnline()
Next

' List Windows Product Activation Status


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colWPA = objWMIService.ExecQuery _
    ("Select * from Win32_WindowsProductActivation")

For Each objWPA in colWPA
    Wscript.Echo "Activation Required: " & objWPA.ActivationRequired
    Wscript.Echo "Description: " & objWPA.Description
    Wscript.Echo "Product ID: " & objWPA.ProductID
    Wscript.Echo "Remaining Evaluation Period: " & _
        objWPA.RemainingEvaluationPeriod
    Wscript.Echo "Remaining Grace Period: " & objWPA.RemainingGracePeriod
    Wscript.Echo "Server Name: " & objWPA.ServerName
Next

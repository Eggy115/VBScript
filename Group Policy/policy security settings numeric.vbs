' Enumerating Resultant Set of Policy Security Settings -- Numeric


strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\rsop\computer")

Set colItems = objWMIService.ExecQuery _
    ("Select * from RSOP_SecuritySettingNumeric")

For Each objItem in colItems
    Wscript.Echo "Key Name: " & objItem.KeyName
    Wscript.Echo "Precedence: " & objItem.Precedence
    Wscript.Echo "Setting: " & objItem.Setting
    Wscript.Echo
Next

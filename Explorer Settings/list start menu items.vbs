
' List Start Menu Items


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_LogicalProgramGroupItem")
 
For Each objItem in colItems
    Wscript.Echo "Name: " & objItem.Name
Next

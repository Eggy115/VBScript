' Identifying Processor Type


strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colProcessors = objWMIService.ExecQuery _
    ("Select * From Win32_Processor")
 
For Each objProcessor in colProcessors
    If objProcessor.Architecture = 0 Then
        Wscript.Echo "This is an x86 computer."
    ElseIf objProcessor.Architecture = 1 Then
        Wscript.Echo "This is a MIPS computer."
    ElseIf objProcessor.Architecture = 2 Then
        Wscript.Echo "This is an Alpha computer."
    ElseIf objProcessor.Architecture = 3 Then
        Wscript.Echo "This is a PowerPC computer."
    ElseIf objProcessor.Architecture = 6 Then
        Wscript.Echo "This is an ia64 computer."
    Else
        Wcript.Echo "The computer type could not be determined."
    End If
Next

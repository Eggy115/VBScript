
' Writing Data to a Text File


Const ForAppending = 8

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    ("c:\scripts\service_status.txt", ForAppending, True)

Set objWMIService=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\cimv2")
Set colServices = objWMIService.ExecQuery("Select * from Win32_Service")

For Each objService in colServices    
    objTextFile.WriteLine(objService.DisplayName & vbTab & _
        objService.State)
Next
objTextFile.Close

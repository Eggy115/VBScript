' List Service Data in a Spreadsheet


Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
objExcel.Workbooks.Add

x = 1
strComputer = "."
Set objWMIService = GetObject _
    ("winmgmts:\\" & strComputer & "\root\cimv2")
Set colServices = objWMIService.ExecQuery _
    ("Select * From Win32_Service")

For Each objService in colServices
    objExcel.Cells(x, 1) = objService.Name
    objExcel.Cells(x, 2) = objService.State
    x = x + 1
Next

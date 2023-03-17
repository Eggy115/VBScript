' AutoFit Columns in a Microsoft Excel Worksheet



Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)
x = 1

strComputer = "."
Set objWMIService = _
    GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
Set colItems = objWMIService.ExecQuery _
    ("Select * from Win32_Service")
For Each objItem in colItems
    objWorksheet.Cells(x, 1) = objItem.Name
    objWorksheet.Cells(x, 2) = objItem.DisplayName
    objWorksheet.Cells(x, 3) = objItem.State
    x = x + 1
Next

Set objRange = objWorksheet.UsedRange
objRange.EntireColumn.Autofit()

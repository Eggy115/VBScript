
' Add Formatted Data to a Spreadsheet



Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True
objExcel.Workbooks.Add
objExcel.Cells(1, 1).Value = "Test value"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 1).Font.Size = 24
objExcel.Cells(1, 1).Font.ColorIndex = 3

' Sort a Microsoft Excel Worksheet



Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add
Set objWorksheet = objWorkbook.Worksheets(1)

objExcel.Cells(1, 1).Value = "4"
objExcel.Cells(2, 1).Value = "1"
objExcel.Cells(3, 1).Value = "2"
objExcel.Cells(4, 1).Value = "3"
objExcel.Cells(1, 2).Value = "A"
objExcel.Cells(2, 2).Value = "B"
objExcel.Cells(3, 2).Value = "C"
objExcel.Cells(4, 2).Value = "D"

Set objRange = objWorksheet.UsedRange
Set objRange2 = objExcel.Range("A1")
objRange.Sort(objRange2)

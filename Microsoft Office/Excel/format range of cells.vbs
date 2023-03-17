' Format a Range of Cells


Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True
objExcel.Workbooks.Add

objExcel.Cells(1, 1).Value = "Name"
objExcel.Cells(1, 1).Font.Bold = TRUE
objExcel.Cells(1, 1).Interior.ColorIndex = 30
objExcel.Cells(1, 1).Font.ColorIndex = 2
objExcel.Cells(2, 1).Value = "Test value 1"
objExcel.Cells(3, 1).Value = "Test value 2"
objExcel.Cells(4, 1).Value = "Tets value 3"
objExcel.Cells(5, 1).Value = "Test value 4"

Set objRange = objExcel.Range("A1","A5")
objRange.Font.Size = 14

Set objRange = objExcel.Range("A2","A5")
objRange.Interior.ColorIndex = 36

Set objRange = objExcel.ActiveCell.EntireColumn
objRange.AutoFit()

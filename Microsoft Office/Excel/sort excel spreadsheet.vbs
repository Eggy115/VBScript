' Sort an Excel Spreadsheet on Three Different Columns



Const xlAscending = 1
Const xlDescending = 2
Const xlYes = 1

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = _ 
    objExcel.Workbooks.Open("C:\Scripts\Sort_test.xls")

Set objWorksheet = objWorkbook.Worksheets(1)
Set objRange = objWorksheet.UsedRange

Set objRange2 = objExcel.Range("A1")
Set objRange3 = objExcel.Range("B1")
Set objRange4 = objExcel.Range("C1")

objRange.Sort objRange2,xlAscending,objRange3,,xlDescending, _
    objRange4,xlDescending,xlYes

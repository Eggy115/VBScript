
' Apply an AutoFormat to an Excel Spreadsheet



Const xpRangeAutoFormatList2 = 11

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Add()
Set objWorksheet = objWorkbook.Worksheets(1)
k = 1
For i = 1 to 10
    For j = 1 to 10
        objWorksheet.Cells(i,j) = k
        k = k + 1
    Next
Next

Set objRange = objWorksheet.UsedRange
objRange.AutoFormat(xpRangeAutoFormatList2)

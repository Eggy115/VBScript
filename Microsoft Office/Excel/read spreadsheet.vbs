' Read an Excel Spreadsheet


Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open _
    ("C:\Scripts\New_users.xls")

intRow = 2

Do Until objExcel.Cells(intRow,1).Value = ""
    Wscript.Echo "CN: " & objExcel.Cells(intRow, 1).Value
    Wscript.Echo "sAMAccountName: " & objExcel.Cells(intRow, 2).Value
    Wscript.Echo "GivenName: " & objExcel.Cells(intRow, 3).Value
    Wscript.Echo "LastName: " & objExcel.Cells(intRow, 4).Value
    intRow = intRow + 1
Loop

objExcel.Quit

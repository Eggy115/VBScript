' Create User Accounts Based on Information in a Spreadsheet


Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open _
    ("C:\Scripts\New_users.xls")

intRow = 2

Do Until objExcel.Cells(intRow,1).Value = ""
    Set objOU = GetObject("ou=Finance, dc=fabrikam, dc=com")
    Set objUser = objOU.Create _
        ("User", "cn=" & objExcel.Cells(intRow, 1).Value)
    objUser.sAMAccountName = objExcel.Cells(intRow, 2).Value
    objUser.GivenName = objExcel.Cells(intRow, 3).Value
    objUser.SN = objExcel.Cells(intRow, 4).Value
    objUser.AccountDisabled = FALSE
    objUser.SetInfo
    intRow = intRow + 1
Loop

objExcel.Quit

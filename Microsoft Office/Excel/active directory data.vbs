' List Active Directory Data in a Spreadsheet


Const ADS_SCOPE_SUBTREE = 2

Set objExcel = CreateObject("Excel.Application")

objExcel.Visible = True
objExcel.Workbooks.Add

objExcel.Cells(1, 1).Value = "Last name"
objExcel.Cells(1, 2).Value = "First name"
objExcel.Cells(1, 3).Value = "Department"
objExcel.Cells(1, 4).Value = "Phone number"

Set objConnection = CreateObject("ADODB.Connection")
Set objCommand =   CreateObject("ADODB.Command")
objConnection.Provider = "ADsDSOObject"
objConnection.Open "Active Directory Provider"

Set objCommand.ActiveConnection = objConnection
objCommand.Properties("Page Size") = 100
objCommand.Properties("Searchscope") = ADS_SCOPE_SUBTREE 
objCommand.CommandText = _
    "SELECT givenName, SN, department, telephoneNumber FROM " _
        & "'LDAP://dc=fabrikam,dc=microsoft,dc=com' WHERE " _
            & "objectCategory='user'"  
Set objRecordSet = objCommand.Execute
objRecordSet.MoveFirst
x = 2

Do Until objRecordSet.EOF
    objExcel.Cells(x, 1).Value = _
        objRecordSet.Fields("SN").Value
    objExcel.Cells(x, 2).Value = _
        objRecordSet.Fields("givenName").Value
    objExcel.Cells(x, 3).Value = _
        objRecordSet.Fields("department").Value
    objExcel.Cells(x, 4).Value = _
        objRecordSet.Fields("telephoneNumber").Value
    x = x + 1
    objRecordSet.MoveNext
Loop

Set objRange = objExcel.Range("A1")
objRange.Activate

Set objRange = objExcel.ActiveCell.EntireColumn
objRange.Autofit()

Set objRange = objExcel.Range("B1")
objRange.Activate
Set objRange = objExcel.ActiveCell.EntireColumn
objRange.Autofit()

Set objRange = objExcel.Range("C1")
objRange.Activate

Set objRange = objExcel.ActiveCell.EntireColumn
objRange.Autofit()

Set objRange = objExcel.Range("D1")
objRange.Activate

Set objRange = objExcel.ActiveCell.EntireColumn
objRange.Autofit()

Set objRange = objExcel.Range("A1").SpecialCells(11)
Set objRange2 = objExcel.Range("C1")
Set objRange3 = objExcel.Range("A1")

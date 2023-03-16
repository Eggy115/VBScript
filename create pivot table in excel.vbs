

' Create a Pivot Table in Excel

Set objExcel = CreateObject("Excel.Application")
objExcel.Visible=True

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

Set xlBook1 = objExcel.WorkBooks.Open(folder & "\pivot.xlsx")

Set objData = xlBook1.Worksheets("data")

Const xlR1C1 = -4150
SrcData = "Data!" & objData.UsedRange.Address(xlR1C1)
'SrcData = "Data!" & objData.Range("A1:D100").Address(xlR1C1)

Set objSheet = xlBook1.Sheets.Add(,objData)
objSheet.Name="Pivot"

Const xlDatabase = 1
Set pvtTable = xlBook1.PivotCaches.Create(xlDatabase,SrcData).CreatePivotTable("Pivot!R1C1","PivotTable1")


Const xlRowField = 1
pvtTable.pivotFields("Name").orientation =xlRowField

Const xlColumnField = 2
pvtTable.pivotFields("Category").orientation = xlColumnField

Const xlFilterField = 3
pvtTable.pivotFields("Date").orientation =  xlFilterField
        
Const xlSum = -4157
pvtTable.AddDataField pvtTable.PivotFields("Value"), "Sum of Value", xlSum

' Read a Fixed-Width Text File Using Database Techniques


On Error Resume Next

Const adOpenStatic = 3
Const adLockOptimistic = 3
Const adCmdText = &H0001

Set objConnection = CreateObject("ADODB.Connection")
Set objRecordSet = CreateObject("ADODB.Recordset")

strPathtoTextFile = "C:\Databases\"

objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
          "Data Source=" & strPathtoTextFile & ";" & _
          "Extended Properties=""text;HDR=YES;FMT=FixedLength"""

objRecordset.Open "SELECT * FROM PhoneList.txt", _
          objConnection, adOpenStatic, adLockOptimistic, adCmdText

Do Until objRecordset.EOF
    Wscript.Echo "Name: " & objRecordset.Fields.Item("FirstName")
    Wscript.Echo "Department: " & objRecordset.Fields.Item("LastName")
    Wscript.Echo "Extension: " & objRecordset.Fields.Item("ID")   
    objRecordset.MoveNext
Loop

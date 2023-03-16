
' Extract Data from Multiple Excel Workbooks

Dim fso
Set fso = CreateObject("Scripting.Filesystemobject")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"


For Each file In fso.GetFolder(folder).Files
  extension=Mid(file.Name,InStrRev(file.Name,"."))

  If extension=".xlsx" Then

    Set objConn = CreateObject("ADODB.Connection")

    objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & file.Path & "; Extended Properties=""Excel 12.0;"";"

    Set ors = objConn.Execute("select * from [data$]")

    Do While Not(ors.EOF)
      WScript.Echo ors(0).Value & " " & ors(1).Value
      ors.MoveNext
    Loop

    ors.Close
    objConn.Close
  End If
Next

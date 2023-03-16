Set objConn = CreateObject("ADODB.Connection")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

objConn.open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & folder & ";Extended Properties=""Text;"";"

Set pivot = CreateObject("Vbsedit.PivotTable")
pivot.Initialize 3,1

Set ors = objConn.Execute("select * from [pivot.csv]")

Do While Not(ors.EOF)
  pivot.Add ors("Date").Value,ors("Name").Value,ors("Category").Value,ors("Value1").Value
  ors.MoveNext
Loop

ors.Close
objConn.Close

pivot.Finalize

pivot.SaveToFile currentdir & "pivot.piv"

For each item1 In pivot.Axe(1)
  Set rowTotal1 = pivot.Aggregate(item1.ID)
  WScript.Echo item1.Label & "[Total:" & rowTotal1.Measure(1) & "]"
  For each item2 In pivot.Axe(2)
    Set rowTotal2 = pivot.Aggregate(item1.ID,item2.ID)
    WScript.Echo "  " & item2.Label & " [Total: " & rowTotal2.Measure(1) & "]"
    For each item3 In pivot.Axe(3)
      Set row = pivot.Aggregate(item1.ID,item2.ID,item3.ID)
      WScript.Echo "    " & row.Label(3) & ": " & row.Measure(1)
    Next
  Next
Next

Set rowTotal = pivot.Aggregate()
WScript.Echo "[Grand Total: " & rowTotal.Measure(1) & "]"

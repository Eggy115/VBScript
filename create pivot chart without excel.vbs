' Create a Pivot Chart Without Excel

Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")

Set shell = CreateObject( "WScript.Shell" )
folder=shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Vbsedit\Resources\"

Set f = fso.OpenTextFile(folder & "pivot.csv",1)
header = f.ReadLine

Set pivot = CreateObject("Vbsedit.PivotTable")
pivot.Initialize 3,2

Do While Not(f.AtEndOfStream)
  line = f.ReadLine
  arr = Split(line,";")
  pivot.Add arr(0),arr(1),arr(2),Replace(arr(3),".",","),Replace(arr(4),".",",")
Loop

pivot.SetColumnNames "Name","Category","Date","Value1","Value2"

pivot.Finalize

pivot.LoadChartTemplate "column"
pivot.ReplaceTag "title","My Column Chart"
pivot.ReplaceTag "bars","vertical"
pivot.ReplaceTag "stacked","false"
pivot.SaveChart folder & "column.htm"

shell.Run folder & "column.htm",1,False

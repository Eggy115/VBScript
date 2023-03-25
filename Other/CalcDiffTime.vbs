
'Check args
Set objArgs = Wscript.Arguments
If objArgs.Count <> 1 then
   WScript.ECHO "Error Args !!"
   WSCript.Quit 1
End If


fromDate=Wscript.Arguments(0)
toDate=FormatDateTime(Now)

wscript.echo DateDiff("s",fromDate,toDate)
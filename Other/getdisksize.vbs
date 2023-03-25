
set oArgs = wscript.arguments
if oArgs.count <> 1 then
	wscript.echo "Error args !!"
	wscript.quit 1
end if

Dim DiskId, objWMIService, objItem, colItems
DiskId = oArgs(0)

Set objWMIService = GetObject ("winmgmts:\\.\root\Microsoft\Windows\Storage")
Set colItems = objWMIService.ExecQuery ("Select * from MSFT_PhysicalDisk where DeviceId=" & DiskId)

For Each objItem in colItems
	' echo size in GB base
	wscript.echo "Size=" & CLng(objItem.Size/1024/1024/1024)
Next

wscript.quit 0
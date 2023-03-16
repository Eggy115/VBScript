' List Fax Server Information


Set objFaxServer = CreateObject("FaxComEx.FaxServer")
objFaxServer.Connect "atl-dc-02"

Wscript.Echo "API version: " & objFaxServer.APIVersion
Wscript.Echo "Major build: " & objFaxServer.MajorBuild
Wscript.Echo "Minor build: " & objFaxServer.MinorBuild
Wscript.Echo "Major version: " & objFaxServer.MajorVersion
Wscript.Echo "Minor version: " & objFaxServer.MinorVersion
Wscript.Echo "Server name: " & objFaxServer.ServerName

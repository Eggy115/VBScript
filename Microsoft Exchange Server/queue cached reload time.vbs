' List the Queued Cache Reload Time


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_QueueCacheReloadEvent")

For Each objItem in colItems
    Wscript.Echo "Reload time: " & objItem.ReloadTime
    Wscript.Echo
Next


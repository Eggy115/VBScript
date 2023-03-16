' List Exchange Domain Controller Information


On Error Resume Next

strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & _
        "\ROOT\MicrosoftExchangeV2")

Set colItems = objWMIService.ExecQuery _
    ("Select * from Exchange_DSAccessDC")

For Each objItem in colItems
    Wscript.Echo "Configuration: " & objItem.Configuration
    Wscript.Echo "Directory type: " & objItem.DirectoryType
    Wscript.Echo "Is fast: " & objItem.IsFast
    Wscript.Echo "Is in sync: " & objItem.IsInSync
    Wscript.Echo "Is up: " & objItem.IsUp
    Wscript.Echo "LDAP port: " & objItem.LDAPPort
    Wscript.Echo "Name: " & objItem.Name
    Wscript.Echo "Type: " & objItem.Type
    Wscript.Echo
Next

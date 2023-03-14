
' Rename a Computer Account


Set objNewOU = GetObject("LDAP://OU=Finance,DC=fabrikam,DC=com")

Set objMoveComputer = objNewOU.MoveHere _
    ("LDAP://CN=atl-pro-037,OU=Finance,DC=fabrikam,DC=com", _
        "CN=atl-pro-003")

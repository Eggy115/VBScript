' Move a Computer Account


Set objNewOU = GetObject("LDAP://OU=Finance,DC=fabrikam,DC=com")

Set objMoveComputer = objNewOU.MoveHere _
    ("LDAP://CN=atl-pro-03,CN=Computers,DC=fabrikam,DC=com", "CN=atl-pro-03")

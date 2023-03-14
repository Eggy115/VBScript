' Reset a Computer Account Password


Set objComputer = GetObject _
    ("LDAP://CN=atl-dc-01,CN=Computers,DC=Reskit,DC=COM")

objComputer.SetPassword "atl-dc-01$"

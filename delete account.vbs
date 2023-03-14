' Delete a Computer Account


strComputer = "atl-pro-040"

set objComputer = GetObject("LDAP://CN=" & strComputer & _
    ",CN=Computers,DC=fabrikam,DC=com")
objComputer.DeleteObject (0)

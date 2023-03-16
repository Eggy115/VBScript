' List User Profile Properties


On Error Resume Next

Set objUser = GetObject _
    ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
 
Wscript.Echo "Profile Path: " & objUser.ProfilePath
Wscript.Echo "Script Path: " & objUser.ScriptPath
Wscript.Echo "Home Directory: " & objUser.HomeDirectory
Wscript.Echo "Home Drive: " & objUser.HomeDrive

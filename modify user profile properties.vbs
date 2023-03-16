' Modify User Profile Properties



Set objUser = GetObject _
  ("LDAP://cn=myerken,ou=management,dc=fabrikam,dc=com")
 
objUser.Put "profilePath", "\\sea-dc-01\Profiles\myerken"
objUser.Put "scriptPath", "logon.bat"
objUser.Put "homeDirectory", "\\sea-dc-01\HomeFolders\myerken"
objUser.Put "homeDrive", "H:"
objUser.SetInfo

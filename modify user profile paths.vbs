' Modify User Profile Paths


Set objUser = GetObject _
    ("LDAP://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
strCurrentProfilePath = objUser.Get("profilePath")
intStringLen = Len(strCurrentProfilePath)
intStringRemains = intStringLen - 11
strRemains = Mid(strCurrentProfilePath, 12, intStringRemains)
strNewProfilePath = "\\fabrikam" & strRemains
objUser.Put "profilePath", strNewProfilePath
objUser.SetInfo

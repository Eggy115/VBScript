' List Object Page Information for a User Account


Set objUser = GetObject _
    ("GC://cn=MyerKen,ou=Management,dc=NA,dc=fabrikam,dc=com")
 
strWhenCreated = objUser.Get("whenCreated")
strWhenChanged = objUser.Get("whenChanged")
 
Set objUSNChanged = objUser.Get("uSNChanged")
dblUSNChanged = _
    Abs(objUSNChanged.HighPart * 2^32 + objUSNChanged.LowPart)
 
Set objUSNCreated = objUser.Get("uSNCreated")
dblUSNCreated = _
    Abs(objUSNCreated.HighPart * 2^32 + objUSNCreated.LowPart)
 
objUser.GetInfoEx Array("canonicalName"), 0
arrCanonicalName = objUser.GetEx("canonicalName")
 
WScript.echo "Canonical Name of object:"
For Each strValue in arrCanonicalName
    WScript.Echo vbTab & strValue
Next
WScript.Echo 
 
WScript.Echo "Object class: " & objUser.Class
WScript.echo "When Created: " & strWhenCreated & " (Created - GMT)"
WScript.echo "When Changed: " & strWhenChanged & " (Modified - GMT)"
WScript.Echo 
WScript.Echo "USN Changed: " & dblUSNChanged & " (USN Current)"
WScript.Echo "USN Created: " & dblUSNCreated & " (USN Original)"

' List Group Object Information


Set objGroup = GetObject _
  ("GC://cn=Scientists,ou=R&D,dc=NA,dc=fabrikam,dc=com")
 
strWhenCreated = objGroup.Get("whenCreated")
strWhenChanged = objGroup.Get("whenChanged")
 
Set objUSNChanged = objGroup.Get("uSNChanged")
dblUSNChanged = _
    Abs(objUSNChanged.HighPart * 2^32 + objUSNChanged.LowPart)
 
Set objUSNCreated = objGroup.Get("uSNCreated")
dblUSNCreated = _
    Abs(objUSNCreated.HighPart * 2^32 + objUSNCreated.LowPart)
 
objGroup.GetInfoEx Array("canonicalName"), 0
arrCanonicalName = objGroup.GetEx("canonicalName")
 
WScript.echo "CanonicalName of object:"
For Each strValue in arrCanonicalName
    WScript.Echo vbTab & strValue
Next
WScript.Echo 
 
WScript.Echo "Object class: " & objGroup.Class 
WScript.Echo "When Created: " & strWhenCreated & " (Created - GMT)"
WScript.Echo "When Changed: " & strWhenChanged & " (Modified - GMT)"
WScript.Echo 
WScript.Echo "USN Changed: " & dblUSNChanged & " (USN Current)"
WScript.Echo "USN Created: " & dblUSNCreated & " (USN Original)"

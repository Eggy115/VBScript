' List the Properties of an OU Object


Set objContainer = GetObject _
   ("GC://ou=Sales,dc=NA,dc=fabrikam,dc=com")
 
strWhenCreated = objContainer.Get("whenCreated")
strWhenChanged = objContainer.Get("whenChanged")
 
Set objUSNChanged = objContainer.Get("uSNChanged")
dblUSNChanged = _
    Abs(objUSNChanged.HighPart * 2^32 + objUSNChanged.LowPart)
 
Set objUSNCreated = objContainer.Get("uSNCreated")
dblUSNCreated = _
    Abs(objUSNCreated.HighPart * 2^32 + objUSNCreated.LowPart)
 
objContainer.GetInfoEx Array("canonicalName"), 0
arrCanonicalName = objContainer.GetEx("canonicalName")
 
WScript.Echo "CanonicalName of object:"
For Each strValue in arrCanonicalName
    WScript.Echo vbTab & strValue
Next
WScript.Echo 
 
WScript.Echo "Object class: " & objContainer.Class & vbCrLf
WScript.Echo "whenCreated: " & strWhenCreated & " (Created - GMT)"
WScript.Echo "whenChanged: " & strWhenChanged & " (Modified - GMT)"
WScript.Echo VbCrLf
WScript.Echo "uSNChanged: " & dblUSNChanged & " (USN Current)"
WScript.Echo "uSNCreated: " & dblUSNCreated & " (USN Original)"

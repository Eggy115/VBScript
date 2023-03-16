' Create 1000 Sample User Accounts


Set objRootDSE = GetObject("LDAP://rootDSE")

Set objContainer = GetObject("LDAP://cn=Users," & _
    objRootDSE.Get("defaultNamingContext"))
 
For i = 1 To 1000
    Set objLeaf = objContainer.Create("User", "cn=UserNo" & i)
    objLeaf.Put "sAMAccountName", "UserNo" & i
    objLeaf.SetInfo
Next
 
WScript.Echo "1000 Users created."

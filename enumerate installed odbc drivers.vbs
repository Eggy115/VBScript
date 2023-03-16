
'Enumerate installed ODBC Drivers on local computer

Const HKEY_LOCAL_MACHINE = &H80000002

Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")
strKeyPath = "SOFTWARE\ODBC\ODBCINST.INI\ODBC Drivers"
objRegistry.EnumValues HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes
For i = 0 to UBound(arrValueNames)
    strValueName = arrValueNames(i)
    objRegistry.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue    
    Wscript.Echo arrValueNames(i) & " — " & strValue
Next

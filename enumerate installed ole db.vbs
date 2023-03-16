
'Enumerate installed OLE DB providers on local computer

Const HKEY_CLASSES_ROOT = &H80000000

Set objRegistry = GetObject("winmgmts:\\.\root\default:StdRegProv")

objRegistry.enumKey HKEY_CLASSES_ROOT, "CLSID", arrKeys

For each key in arrKeys
  If objRegistry.GetDWordValue (HKEY_CLASSES_ROOT,"CLSID\" & key,"OLEDB_SERVICES",uValue)=0 Then
    objRegistry.GetStringValue HKEY_CLASSES_ROOT,"CLSID\" & key,"",providerName
    objRegistry.GetStringValue HKEY_CLASSES_ROOT,"CLSID\" & key & "\OLE DB Provider","",description
    WScript.Echo "[" & providerName & "] " & description
  End if
Next

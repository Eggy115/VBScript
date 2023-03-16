
' Modify the ODBC Default Value to Comma-Delimited When Reading Text Files


Const HKEY_LOCAL_MACHINE = &H80000002

strComputer = "."
Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv")

strKeyPath = "SOFTWARE\Microsoft\Jet\4.0\Engines\Text"
strValueName = "Format"
strValue = "CSVDelimited"
objReg.SetStringValue _
    HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue

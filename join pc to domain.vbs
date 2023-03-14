' Join a Computer to a Domain


Const JOIN_DOMAIN = 1
Const ACCT_CREATE = 2
Const ACCT_DELETE = 4
Const WIN9X_UPGRADE = 16
Const DOMAIN_JOIN_IF_JOINED = 32
Const JOIN_UNSECURE = 64
Const MACHINE_PASSWORD_PASSED = 128
Const DEFERRED_SPN_SET = 256
Const INSTALL_INVOCATION = 262144
 
strDomain = "FABRIKAM"
strPassword = "ls4k5ywA"
strUser = "shenalan"
 
Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName
 
Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & _
    strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & _
        strComputer & "'")
 
ReturnValue = objComputer.JoinDomainOrWorkGroup(strDomain, _
    strPassword, strDomain & "\" & strUser, NULL, _
        JOIN_DOMAIN + ACCT_CREATE)

' Unjoin a Computer from a Domain



Const NETSETUP_ACCT_DELETE = 2 'Disables computer account in domain.
strPassword = "password"
strUser = "kenmyer"

Set objNetwork = CreateObject("WScript.Network")
strComputer = objNetwork.ComputerName

Set objComputer = GetObject("winmgmts:{impersonationLevel=Impersonate}!\\" & _
 strComputer & "\root\cimv2:Win32_ComputerSystem.Name='" & strComputer & "'")
strDomain = objComputer.Domain
intReturn = objComputer.UnjoinDomainOrWorkgroup _
 (strPassword, strDomain & "\" & strUser, NETSETUP_ACCT_DELETE)

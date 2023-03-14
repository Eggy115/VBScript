' List FSMO Role Holders


Set objRootDSE = GetObject("LDAP://rootDSE")
 
Set objSchema = GetObject _
    ("LDAP://" & objRootDSE.Get("schemaNamingContext"))
strSchemaMaster = objSchema.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strSchemaMaster)
Set objComputer = GetObject(objNtds.Parent)
WScript.Echo "Forest-wide Schema Master FSMO: " & objComputer.Name
 
Set objNtds = Nothing
Set objComputer = Nothing
 
Set objPartitions = GetObject("LDAP://CN=Partitions," & _ 
    objRootDSE.Get("configurationNamingContext"))
strDomainNamingMaster = objPartitions.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strDomainNamingMaster)
Set objComputer = GetObject(objNtds.Parent)
WScript.Echo "Forest-wide Domain Naming Master FSMO: " & objComputer.Name
 
Set objDomain = GetObject _
    ("LDAP://" & objRootDSE.Get("defaultNamingContext"))
strPdcEmulator = objDomain.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strPdcEmulator)
Set objComputer = GetObject(objNtds.Parent)
WScript.Echo "Domain's PDC Emulator FSMO: " & objComputer.Name
 
Set objRidManager = GetObject("LDAP://CN=RID Manager$,CN=System," & _
    objRootDSE.Get("defaultNamingContext"))
strRidMaster = objRidManager.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strRidMaster)
Set objComputer = GetObject(objNtds.Parent)
WScript.Echo "Domain's RID Master FSMO: " & objComputer.Name
 
Set objInfrastructure = GetObject("LDAP://CN=Infrastructure," & _
    objRootDSE.Get("defaultNamingContext"))
strInfrastructureMaster = objInfrastructure.Get("fSMORoleOwner")
Set objNtds = GetObject("LDAP://" & strInfrastructureMaster)
Set objComputer = GetObject(objNtds.Parent)
WScript.Echo "Domain's Infrastructure Master FSMO: " & objComputer.Name

' Convert domain name to ADsPath



strDomainName = "accounting.sea.na.fabrikam.com"
arrDomLevels = Split(strDomainName, ".")
strADsPath = "dc=" & Join(arrDomLevels, ",dc=")
WScript.Echo strADsPath

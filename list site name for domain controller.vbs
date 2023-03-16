
' List the Site Name for a  Domain Controller


strDcName = "atl-dc-01"
Set objADSysInfo = CreateObject("ADSystemInfo")

strDcSiteName = objADSysInfo.GetDCSiteName(strDcName)
WScript.Echo "DC Site Name: " & strDcSiteName

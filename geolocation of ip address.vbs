
' Geolocation of IP Address - IP to Country

Set ip = CreateObject("Vbsedit.iptocountry")

Set fso = CreateObject("Scripting.Filesystemobject")

If Not(fso.FileExists("iptocountry.bin")) Then
  If fso.FileExists("iptocountry.csv") Then
    ip.CreateDatabaseFromCsv "iptocountry.csv","iptocountry.bin"
  Else
    WScript.Quit
  End If
Else
  If fso.FileExists("iptocountry.csv") Then
    If fso.GetFile("iptocountry.csv").DateLastModified > _
          fso.GetFile("iptocountry.bin").DateLastModified Then
      ip.CreateDatabaseFromCsv "iptocountry.csv","iptocountry.bin"
    End If
  End If
End If

'if you need to convert only a few IP addresses
ip.DatabasePath = "iptocountry.bin"

'if you need to convert a lot of IP addresses
'ip.LoadDatabaseIntoMemory "iptocountry.bin"

code = ip.IpToCountry("185.238.208.18")
WScript.Echo code

WScript.Echo ip.CountryCodeToLongCode(code)
WScript.Echo ip.CountryCodeToCountryName(code)
